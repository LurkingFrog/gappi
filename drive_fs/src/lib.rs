use log::{debug, info};
use std::cell::RefCell;
use std::collections::HashMap;

pub use wrapi::{AuthMethod, WrapiApi, WrapiError, WrapiResult};
pub mod models;

#[derive(Clone, Debug)]
pub struct FileNode {
  id: String,
  name: String,
  parents: Vec<String>,
  children: Vec<String>,
}

#[derive(Clone, Debug)]
pub struct FileCache {
  root_id: String,
  // Organize the files by parents/children for easy traversal to root
  graph_cache: HashMap<String, FileNode>,
  // Lookup a File by its path, which can be calculated from the graph
  path_cache: HashMap<String, String>,
}

impl FileCache {
  fn empty() -> FileCache {
    FileCache {
      root_id: "root".to_string(),
      graph_cache: HashMap::new(),
      path_cache: HashMap::new(),
    }
  }

  // TODO: add clear cache function

  /// Get all the directories loaded into the cache so we can do a quick find
  fn load(&self, api: std::cell::RefMut<impl wrapi::WrapiApi>) -> Result<FileCache, WrapiError> {
    info!("Loading the cache");
    let request = models::FileRequest {
      parent_id: "root".to_string(),
      filters: vec![models::FileFilter::Type(models::MimeType::Folder)],
      opts: vec![],
    };
    let result: Box<models::FileResult> = api.call("find", request)?;
    let mut graph: HashMap<String, FileNode> = HashMap::new();

    // A disposable hash to find the root node (since it has an ID, but does not show up in the query)
    // The root will be the only entry with a parent count of 0
    let mut root_finder: HashMap<String, usize> = HashMap::new();
    for file in result.files {
      let id = file.id.unwrap();
      let name = file.name.unwrap().clone();

      let parents = match file.parents {
        Some(parents) => parents.clone(),
        None => vec![],
      };

      // Load the file's info into the root finder
      for parent in parents.clone() {
        root_finder.entry(parent.clone()).or_insert(0);

        // Add the child to the parent, initializing it if it doesn't exist
        graph
          .entry(parent.clone())
          .and_modify({ |node| node.children.push(id.clone()) })
          .or_insert(FileNode {
            id: parent.clone(),
            name: "Unknown: Parent Insert".to_string(),
            parents: vec![],
            children: vec![id.clone()],
          });
      }
      root_finder
        .entry(id.clone())
        .and_modify(|e| *e = parents.len())
        .or_insert(parents.len());

      // And add it to the graph
      graph
        .entry(id.clone())
        .and_modify({
          |node| {
            *node = FileNode {
              id: id.clone(),
              name: name.clone(),
              parents: parents.clone(),
              children: node.children.clone(),
            }
          }
        })
        .or_insert(FileNode {
          id: id.clone(),
          name: name.clone(),
          parents: parents.clone(),
          children: vec![],
        });
    }

    let root: Option<String> = root_finder.keys().into_iter().fold(Ok(None), |acc, key| {
      let count = root_finder.get(key).unwrap();
      match (acc, count) {
        (Ok(None), 0) => Ok(Some(key.clone())),
        (Ok(Some(x)), 0) => Err(format!(
          "Found multiple values without parents: ({}, {})",
          x, key
        )),
        (err, _) => err,
      }
    })?;

    let root_id = match root {
      Some(x) => x,
      None => Err("Did not find a root path without parents")?,
    };

    // Unravel the graph
    // If root in id, add to children
    // create root node
    // add back to graph
    let root_node = graph.clone().into_iter().fold(
      FileNode {
        id: root_id.clone(),
        name: "root".to_string(),
        parents: vec![],
        children: vec![],
      },
      |acc, (key, value)| match value.parents.contains(&root_id) {
        true => {
          let mut children = acc.children.clone();
          children.push(key);
          FileNode {
            id: acc.id,
            name: value.name.clone(),
            children: children,
            parents: acc.parents,
          }
        }
        false => acc,
      },
    );

    let mut path_map = HashMap::new();
    fn path_builder(
      cwd: String,
      file: FileNode,
      graph: &HashMap<String, FileNode>,
      path_map: &mut HashMap<String, String>,
    ) -> Result<(), WrapiError> {
      path_map.entry(cwd.clone()).or_insert(file.id.clone());
      match file.children.is_empty() {
        true => Ok(()),
        false => {
          for child in file.children.clone() {
            match graph.get(&child) {
              None => Err(format!(
                "Path builder - could not find child '{}' in graph",
                child
              ))?,
              Some(node) => path_builder(
                match cwd.len() {
                  0 | 1 => format!("/{}", node.name),
                  _ => format!("{}/{}", cwd, node.name),
                },
                node.clone(),
                graph,
                path_map,
              )?,
            }
          }
          Ok(())
        }
      }
    }
    path_builder("/".to_string(), root_node, &graph, &mut path_map)?;
    // let dirs =
    // If Parent exists, append to vec

    // println!("path_map:\n{:#?}", path_map);
    // Load them into the path cache
    Ok(FileCache {
      root_id: root_id.clone(),
      graph_cache: graph,
      path_cache: path_map,
    })
  }
}

/// A struct to contain the API and link all the calls to
#[derive(Debug)]
pub struct DriveFS {
  api: RefCell<wrapi::API>,
  cache: FileCache,
}

impl DriveFS {
  pub fn build(auth: wrapi::AuthMethod) -> DriveFS {
    let api = wrapi::API::new(auth.clone()).add_endpoint(
      "find".to_string(),
      wrapi::Endpoint {
        base_url: "https://www.googleapis.com/drive/v3/files",
        auth_method: auth.clone(),
        request_method: wrapi::RequestMethod::GET,
        scopes: vec!["https://www.googleapis.com/auth/drive"],
        request_mime_type: wrapi::MimeType::Json,
        response_mime_type: wrapi::MimeType::Json,
      },
    );
    let mut path_cache = HashMap::new();
    path_cache.insert("/", "root".to_string());
    path_cache.insert("", "root".to_string());

    DriveFS {
      api: RefCell::new(api),
      cache: FileCache::empty(),
    }
  }

  pub fn load_cache(self) -> Result<DriveFS, WrapiError> {
    let new_cache = self.cache.load(self.api.borrow_mut())?;
    Ok(DriveFS {
      api: self.api,
      cache: new_cache,
    })
  }

  // TODO: Simplify this. With the path cache, this should be a direct lookup but seems to be looking
  fn get_path_id(&self, path: &str) -> Result<String, WrapiError> {
    println!("Finding the ID for directory:\n{:#?}", path);
    let parent_id = path.split("/").into_iter().fold(
      Ok(("root".to_string(), "/".to_string())),
      |acc: Result<(String, String), WrapiError>, x| {
        let (current_id, cwd) = acc?;
        match (cwd.as_ref(), x) {
          ("/", "") => Ok((current_id, String::from(""))),
          (_, "") => Err(WrapiError::Json(format!(
            "In fold and found a double '/' in the path: {}",
            x
          ))),
          (_, _) => {
            let next_dir = format!("{}/{}", cwd, x);
            match self.cache.path_cache.get(&next_dir[..]) {
              Some(next_id) => Ok((next_id.clone(), next_dir)),
              None => {
                println!("{} not found in cache. Doing the lookup now", next_dir);
                let request = models::FileRequest {
                  parent_id: current_id.clone(),
                  filters: vec![],
                  opts: vec![],
                };
                let _result: Box<models::FileResult> =
                  self.api.borrow_mut().call("find", request)?;
                Ok(("NOT ROOT".to_string(), next_dir))
              }
            }
          }
        }
      },
    )?;
    Ok(parent_id.0)
  }

  pub fn ls(
    &self,
    path: &str,
    _opts: Vec<models::FileOpts>,
  ) -> Result<models::FileResult, WrapiError> {
    // Recursive calls on the path to find the right one.
    let pwd = self.get_path_id(path)?;
    Err(WrapiError::General(format!(
      "Exiting LS: Found PWD of {}",
      pwd
    )))
  }

  // TODO: Parent ID does not work. Make test to validate and verify
  pub fn find(
    self,
    work_dir: &str,
    filters: Vec<models::FileFilter>,
    opts: Vec<models::FileOpts>,
  ) -> Result<Box<models::FileResult>, WrapiError> {
    let parent_id = self.get_path_id(work_dir)?;
    debug!("parent_id:\n{:#?}", parent_id);

    // TODO: LS each sub-directory
    // Get all the files with the name
    let request = models::FileRequest {
      parent_id: parent_id,
      filters: filters,
      opts: opts,
    };
    let result = self.api.borrow_mut().call("find", request)?;

    Ok(result)
  }

  // pub fn mkdir(&self, path: &str, opts: ) {}
}
