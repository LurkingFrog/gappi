use log::{debug, info};

use drive_fs::{models, DriveFS};

#[test]
fn test_connection() {
  env_logger::init();

  info!("In drive_fs::test_connection");
  let service_account = wrapi::build_service_account(
    "../test_data/service_acct.json".to_string(),
    Some("fhl@landfillinc.com".to_string()),
  );
  let drive_api = DriveFS::build(service_account)
    .load_cache()
    .expect("Error loading the cache");

  let files = drive_api
    .find(
      "/Sandbox",
      vec![
        models::FileFilter::Name(models::Filter::Equals("Submissions_Log".to_string())),
        models::FileFilter::Type(models::MimeType::Spreadsheet),
      ],
      vec![models::FileOpts::IsUnique(true)],
    )
    .unwrap();
  debug!("files:\n{:#?}", files);
}
