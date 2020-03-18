use log::info;

#[test]
fn test_connection() {
  env_logger::init();

  info!("In sheets_db::test_connection");
  let service_account = wrapi::build_service_account(
    "../test_data/service_acct.json".to_string(),
    Some("fhl@landfillinc.com".to_string()),
  );

  let sheet_id = "1kwQgjicMgKVV1aZ1oStIjpahQLDronaqzkTKdD-paI0".to_string();
  let worksheet =
    sheets_db::SheetDB::open(service_account, sheet_id).expect("Error opening the worksheet");

  let sheets = worksheet.list_sheets();
  info!("Listed Sheets:\n{:#?}", sheets);

  // Define sheets api call

  // Simple call to sheets
}
