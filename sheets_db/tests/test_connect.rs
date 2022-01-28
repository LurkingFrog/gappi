use log::{debug, error, info};

use std::sync::Once;
static INIT: Once = Once::new();
pub fn test_init() {
  println!("\n\n\n");
  INIT.call_once(|| env_logger::init());
}

#[test]
fn test_connection() {
  test_init();

  info!("In sheets_db::test_connection");
  let service_account = wrapi::build_service_account(
    // "../test_data/service_acct.json".to_string(),
    "/home/dfogelson/.secure/fhl_service_acct.json".to_string(),
    Some("fhl@landfillinc.com".to_string()),
  );

  let sheet_id = "1kwQgjicMgKVV1aZ1oStIjpahQLDronaqzkTKdD-paI0".to_string();
  let worksheet =
    sheets_db::SheetDB::open(service_account, sheet_id).expect("Error opening the worksheet");

  let _sheets = worksheet.list_sheets();

  // Define sheets api call

  // Simple call to sheets
}

#[test]
fn test_json() {
  test_init();

  info!("Testing a Json Block");
  let value = std::fs::read_to_string("../tmp.json");
  let result: sheets_db::Spreadsheet = match value {
    Ok(val) => {
      // debug! {"Sheet Data:\n{:#?}", val};
      serde_json::from_str(&val).unwrap()
    }
    Err(err) => {
      error!("Couldn't read file: {:?}", err);
      panic!("Exiting now");
    }
  };
  debug!("{:?}", result);
}
