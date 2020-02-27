use serde_derive::{Deserialize, Serialize};
use wrapi::{WrapiError, WrapiRequest, WrapiResult};

#[derive(Clone, Debug, Serialize, Deserialize)]
// #[serde(untagged)]
pub enum MimeType {
  #[serde(rename = "application/vnd.google-apps.audio")]
  Audio,
  #[serde(rename = "text/csv")]
  CSV,
  #[serde(rename = "application/vnd.google-apps.document")] //	Google Docs
  Doc,
  #[serde(rename = "application/vnd.google-apps.drawing")] //	Google Drawing
  Drawing,
  #[serde(rename = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]
  Excel,
  #[serde(rename = "application/vnd.google-apps.file")] //	Google Drive file
  File,
  #[serde(rename = "application/vnd.google-apps.folder")] //	Google Drive folder
  Folder,
  #[serde(rename = "application/vnd.google-apps.form")] //	Google Forms
  Form,
  #[serde(rename = "application/vnd.google-apps.fusiontable")] //	Google Fusion Tables
  FusionTable,
  #[serde(rename = "image/gif")]
  Gif,
  #[serde(rename = "text/html")] //	Html Document
  HTML,
  #[serde(rename = "application/vnd.google-apps.map")] //	Google My Maps
  Map,
  #[serde(rename = "image/jpeg")]
  Jpeg,
  #[serde(rename = "application/pdf")]
  PDF,
  #[serde(rename = "application/vnd.google-apps.photo")] //
  Photo,
  #[serde(rename = "image/png")]
  PNG,
  #[serde(rename = "application/vnd.google-apps.presentation")] //	Google Slides
  Presentation,
  #[serde(rename = "application/vnd.google-apps.script")] //	Google Apps Scripts
  Script,
  #[serde(rename = "application/vnd.google-apps.drive-sdk")] //	3rd party shortcut
  Shortcut,
  #[serde(rename = "application/vnd.google-apps.site")] //	Google Sites
  Site,
  #[serde(rename = "application/vnd.google-apps.spreadsheet")] //	Google Sheets
  Spreadsheet,
  #[serde(rename = "text/plain")]
  Text,
  #[serde(rename = "application/vnd.google-apps.unknown")]
  Unknown,
  #[serde(rename = "application/vnd.google-apps.video")] //
  Video,
  #[serde(rename = "application/vnd.openxmlformats-officedocument.wordprocessingml.document")] //
  Word,
  #[serde(rename = "application/zip")]
  Zip,
}

impl MimeType {
  pub fn to_string(&self) -> String {
    let mut value = serde_json::to_string(self).unwrap();
    value.pop();
    value.remove(0);
    value
  }
}

// impl<'de> serde::Deserialize<'de> for Kind {
//     fn deserialize<D>(des: D) -> Result<Self, D::Error>
//     where
//         D: serde::Deserializer<'de>,
//     {
//     }
// }

#[derive(Default, Clone, Debug, Serialize, Deserialize)]
pub struct File {
  /// The MIME type of the file.
  /// Google Drive will attempt to automatically detect an appropriate value from uploaded content if no value is provided. The value cannot be changed unless a new revision is uploaded.
  /// If a file is created with a Google Doc MIME type, the uploaded content will be imported if possible. The supported import formats are published in the About resource.
  #[serde(rename = "mimeType")]
  pub mime_type: Option<MimeType>,
  /// The ID of the file.
  pub id: Option<String>,
  pub parents: Option<Vec<String>>,
  /// Links for exporting Google Docs to specific formats.
  #[serde(rename = "exportLinks")]
  /// A short description of the file.
  pub description: Option<String>,
  /// Identifies what kind of resource this is. Value: the fixed string "drive#file".
  pub kind: Option<String>,
  /// The name of the file. This is not necessarily unique within a folder. Note that for immutable items such as the top level folders of shared drives, My Drive root folder, and Application Data folder the name is constant.
  pub name: Option<String>,
  /// A link for downloading the content of the file in a browser. This is only available for files with binary content in Google Drive.
  #[serde(rename = "webContentLink")]
  pub web_content_link: Option<String>,
  /// ID of the shared drive the file resides in. Only populated for items in shared drives.
  #[serde(rename = "driveId")]
  pub drive_id: Option<String>,
  /// The list of spaces which contain the file. The currently supported values are 'drive', 'appDataFolder' and 'photos'.
  pub spaces: Option<Vec<String>>,
  /// Whether the file has been trashed, either explicitly or from a trashed parent folder. Only the owner may trash a file, and other users cannot see files in the owner's trash.
  pub trashed: Option<bool>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum Kind {
  #[serde(rename = "drive#fileList")]
  FileList,
  #[serde(rename = "drive#file")]
  File(File),
  None,
}

impl Default for Kind {
  fn default() -> Self {
    Kind::None
  }
}

#[derive(Default, Clone, Debug, Serialize, Deserialize)]
/// If we run a query, this contains the results and has some functions for dealing with them
pub struct FileSearchResult {
  pub kind: Kind,
  #[serde(rename = "incompleteSearch")]
  pub incomplete_search: bool,
  pub files: Vec<File>,
}

impl FileSearchResult {}

#[derive(Default, Clone, Debug, Serialize, Deserialize)]
pub struct FileMetadata {
  name: String,
  #[serde(rename = "mimeType")]
  mime_type: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct CreateFolder {
  /// If a file is created with a Google Doc MIME type, the uploaded content will be imported if possible. The supported import formats are published in the About resource.
  #[serde(rename = "mimeType")]
  pub mime_type: MimeType,
  /// The name of the file. This is not necessarily unique within a folder. Note that for immutable items such as the top level folders of shared drives, My Drive root folder, and Application Data folder the name is constant.
  pub name: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct CreateFile {
  /// If a file is created with a Google Doc MIME type, the uploaded content will be imported if possible. The supported import formats are published in the About resource.
  #[serde(rename = "mimeType")]
  pub mime_type: MimeType,
  /// The name of the file. This is not necessarily unique within a folder. Note that for immutable items such as the top level folders of shared drives, My Drive root folder, and Application Data folder the name is constant.
  pub name: String,
  // The ID of the parent
  pub parent: String,
}

// ******************************************
// *****                                *****
// *****              Calls             *****
// *****                                *****
// ******************************************
#[derive(Clone, Debug)]
pub enum Filter {
  And(Box<(Filter, Filter)>),
  Or(Box<(Filter, Filter)>),
  Not(Box<Filter>),
  Any(Vec<Filter>),
  All(Vec<Filter>),
  Contains(String),
  Equals(String),
  StartsWith(String),
  DateGT(String), // This should be a Chrono object
  DateLT(String),
}

impl Filter {
  pub fn to_string(&self) -> Result<String, WrapiError> {
    match self {
      Filter::Equals(value) => Ok(format!("= '{}'", value)),
      _ => Err(WrapiError::Json(
        "Filter option not implemented".to_string(),
      )),
    }
  }
}

// TODO: Do we need to make a back reference to the filter for Not, Any, Or?
#[derive(Clone, Debug)]
pub enum FileFilter {
  Type(MimeType),
  Name(Filter),
  FullText(Filter),
  Parent(Filter), // TODO: Can this be context sensitive, rejecting Contains?
  IsSharedWithMe(bool),
}

impl FileFilter {
  pub fn to_string(&self) -> Result<String, WrapiError> {
    match self {
      FileFilter::Type(mime_type) => Ok(format!("mimeType = '{}'", mime_type.to_string())),
      FileFilter::Name(filter) => Ok(format!("name {}", filter.to_string()?)),
      _ => Err(WrapiError::Json(
        "FileFilter option not implemented".to_string(),
      )),
    }
  }
}

pub enum FileOpts {
  /// Include files that were deleted
  // default: false
  IncludeTrashed(bool),
  /// Search recursively for anything that matches the filter
  Recursive(bool),
  /// Throw an error if any count other than one is found
  IsUnique(bool),
}

pub struct FileRequest {
  pub parent_id: String,
  pub filters: Vec<FileFilter>,
  pub opts: Vec<FileOpts>,
}

impl WrapiRequest for FileRequest {
  fn build_uri(&self, base_url: &str) -> Result<String, WrapiError> {
    let mut query_params: Vec<String> = vec![];
    for filter in &self.filters {
      query_params.push(filter.to_string()?)
    }
    query_params.push(String::from("trashed=false"));
    let query = query_params.join(" and ");
    println!("Query String: {}", query);
    let uri = url::Url::parse_with_params(
      base_url,
      &[
        ("q", &query[..]),
        (
          "fields",
          "kind,nextPageToken,incompleteSearch,files/id,files/name,files/mimeType,files/parents",
        ),
      ],
    )
    .unwrap()
    .into_string()
    .parse()
    .unwrap();

    Ok(uri)
  }

  fn build_body(&self) -> Result<String, WrapiError> {
    Ok("".to_string())
  }

  fn build_headers(&self) -> Result<Vec<(String, String)>, WrapiError> {
    Ok(vec![])
  }
}

#[derive(Debug, Serialize, Deserialize)]
pub struct FileResult {
  pub files: Vec<File>,
}

impl WrapiResult for FileResult {
  fn parse(_headers: Vec<(String, String)>, body: Vec<u8>) -> Result<Box<FileResult>, WrapiError> {
    // println!("Serde Result:\n{:#?}", std::str::from_utf8(&b ody));
    let result: FileSearchResult = serde_json::from_str(std::str::from_utf8(&body)?)?;
    Ok(Box::new(FileResult {
      files: result.files,
    }))
  }
}
