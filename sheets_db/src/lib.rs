/// Functional interactios with Google Sheets (V4)
use log::{debug, error, info};
use std::cell::RefCell;
use std::collections::HashMap;

use serde_derive::{Deserialize, Serialize};
use wrapi::{WrapiApi, WrapiError, WrapiRequest, WrapiResult};

// Object Regex:  \s*\{\s*object \((\w+)\)\s*\}
// Vec Object Regex: \[\s*\{\s*object \((\w+)\)\s*\}\s*\]
// De-Property Regex: "(\w+)"\s*:

/// Default function to set a f32 value to 1
fn one() -> f32 {
  1.0
}

fn bool_false() -> bool {
  false
}

// #[derive(Clone, Debug, Serialize, Deserialize)]
// pub enum Value {
//   #[serde(rename = "null_value")]
//   Null,
//   #[serde(rename = "number_value")]
//   Number,
//   #[serde(rename = "string_value")]
//   Str,
//   #[serde(rename = "bool_value")]
//   Bool,
//   #[serde(rename = "struct")]
//   Obj,
//   #[serde(rename = "ListValue")]
//   ListValue(Vec<Value>),
// }

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum RecalculationInterval {
  //	Default value. This value must not be used.
  #[serde(rename = "RECALCULATION_INTERVAL_UNSPECIFIED")]
  RecalculationIntervalUnspecified,
  //	Volatile functions are updated on every change.
  #[serde(rename = "ON_CHANGE")]
  OnChange,
  //	Volatile functions are updated on every change and every minute.
  #[serde(rename = "MINUTE")]
  Minute,
  #[serde(rename = "HOUR")]
  Hour,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum NumberFormatType {
  // The number format is not specified and is based on the contents of the cell. Do not explicitly use this.
  #[serde(rename = "NUMBER_FORMAT_TYPE_UNSPECIFIED")]
  Unspecified,
  #[serde(rename = "TEXT")]
  Text,
  #[serde(rename = "NUMBER")]
  Number,
  #[serde(rename = "PERCENT")]
  Percent,
  #[serde(rename = "CURRENCY")]
  Currency,
  #[serde(rename = "DATE")]
  Date,
  #[serde(rename = "TIME")]
  Time,
  #[serde(rename = "DATE_TIME")]
  DateTime,
  #[serde(rename = "SCIENTIFIC")]
  Scientific,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct NumberFormat {
  #[serde(rename = "type")]
  format_type: NumberFormatType,
  pattern: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct Color {
  red: f32,
  green: f32,
  blue: f32,
  #[serde(default = "one")]
  alpha: f32,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct Borders {
  borders_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct Padding {
  top: i16,
  right: i16,
  bottom: i16,
  left: i16,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum HorizontalAlign {
  Horizontalalign_not_implemented,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum VertialAlign {
  VERTIALALIGN_NOT_IMPLEMENTED,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum WrapStrategy {
  #[serde(rename = "")]
  WRAPSTRATEGY_NOT_IMPLEMENTED,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum TextDirection {
  #[serde(rename = "")]
  TEXTDIRECTION_NOT_IMPLEMENTED,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum HyperlinkDisplayType {
  #[serde(rename = "")]
  HYPERLINKDISPLAYTYPE_NOT_IMPLEMENTED,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct VerticalAlign {
  verticalalign_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct TextFormat {
  textformat_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct TextRotation {
  textrotation_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct CellFormat {
  #[serde(rename = "numberFormat")]
  number_format: NumberFormat,
  #[serde(rename = "backgroundColor")]
  background_color: Color,
  borders: Borders,
  padding: Padding,
  #[serde(rename = "horizontalAlignment")]
  horizontal_alignment: HorizontalAlign,
  #[serde(rename = "verticalAlignment")]
  vertical_alignment: VerticalAlign,
  #[serde(rename = "wrapStrategy")]
  wrap_strategy: WrapStrategy,
  #[serde(rename = "textDirection")]
  text_direction: TextDirection,
  #[serde(rename = "textFormat")]
  text_format: TextFormat,
  #[serde(rename = "hyperlinkDisplayType")]
  hyperlink_display_type: HyperlinkDisplayType,
  #[serde(rename = "textRotation")]
  text_rotation: TextRotation,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct IterativeCalculationSettings {
  iterativecalculationsettings_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct SpreadsheetTheme {
  spreadsheettheme_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct SpreadsheetProperties {
  title: String,
  locale: String,
  #[serde(rename = "autoRecalc")]
  auto_recalc: RecalculationInterval,
  #[serde(rename = "timeZone")]
  timezone: String,
  #[serde(rename = "defaultFormat")]
  default_format: CellFormat,
  #[serde(rename = "iterativeCalculationSettings")]
  iterative_calculation_settings: IterativeCalculationSettings,
  #[serde(rename = "spreadsheetTheme")]
  spreadsheet_theme: SpreadsheetTheme,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct GridProperties {
  #[serde(rename = "rowCount")]
  pub row_count: i32,
  #[serde(rename = "columnCount")]
  pub column_count: i32,
  #[serde(default, rename = "frozenRowCount")]
  frozen_row_count: i32,
  #[serde(default, rename = "frozenColumnCount")]
  frozen_column_count: i32,
  #[serde(rename = "hideGridLines", default = "bool_false")]
  hide_grid_lines: bool,
  #[serde(default = "bool_false", rename = "rowGroupControlAfter")]
  row_group_control_after: bool,
  #[serde(default = "bool_false", rename = "columnGroupControlAfter")]
  column_group_control_after: bool,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum SheetType {
  #[serde(rename = "SHEET_TYPE_UNSPECIFIED")]
  Unspecified,
  #[serde(rename = "GRID")]
  Grid,
  #[serde(rename = "OBJECT")]
  Object,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct SheetProperties {
  #[serde(rename = "sheetId")]
  sheet_id: i64,
  title: String,
  index: i32,
  #[serde(rename = "sheetType")]
  sheet_type: SheetType,
  #[serde(rename = "gridProperties")]
  pub grid_properties: GridProperties,
  #[serde(default = "bool_false")]
  hidden: bool,
  #[serde(rename = "tabColor")]
  tab_color: Option<Color>,
  #[serde(default = "bool_false", rename = "rightToLeft")]
  right_to_left: bool,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct GridData {
  griddata_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct GridRange {
  #[serde(rename = "sheetId")]
  sheet_id: i64,
  /// The is my own addition. Hopefully it won't break things
  // sheet_name: Option<String>,
  #[serde(rename = "startRowIndex")]
  start_row_index: i32,
  #[serde(rename = "endRowIndex")]
  end_row_index: i32,
  #[serde(rename = "startColumnIndex")]
  start_column_index: i32,
  #[serde(rename = "endColumnIndex")]
  end_column_index: i32,
}

impl GridRange {
  // Convert column index into the letters that Excel/Sheets uses for naming
  pub fn to_base_26(value: i32) -> String {
    let mut result = "".to_string();
    let mut remainder = value.clone() as usize - 1;
    let alpha_map = [
      "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R",
      "S", "T", "U", "V", "W", "X", "Y", "Z",
    ];

    let mut flag = true;
    while flag {
      match remainder / 26 {
        0 => {
          result = format!("{}{}", result, alpha_map[remainder].clone());
          flag = false;
        }
        index => {
          result = format!("{}{}", result, alpha_map[index - 1].clone());
          remainder %= 26;
        }
      }
    }
    result
  }

  pub fn to_range(&self, sheet_name: String) -> String {
    format!(
      "'{}'!{}{}:{}{}",
      sheet_name,
      GridRange::to_base_26(self.start_column_index),
      self.start_row_index,
      GridRange::to_base_26(self.end_column_index),
      self.end_row_index,
    )
  }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum MajorDimension {
  #[serde(rename = "DIMENSION_UNSPECIFIED")]
  DimensionUnspecified,
  #[serde(rename = "ROWS")]
  Rows,
  #[serde(rename = "COLUMNS")]
  Columns,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum Value {
  Null,
  Number(f64),
  StringValue(String),
  Bool(bool),
  Struct(HashMap<String, Value>),
  List(Vec<Value>),
}

impl IntoIterator for Value {
  type Item = Value;
  type IntoIter = std::vec::IntoIter<Self::Item>;

  fn into_iter(self) -> Self::IntoIter {
    match self {
      Value::List(values) => values.clone().into_iter(),
      _ => panic!("Tried to convert a non-list sheets_db::Value into an iterator"),
    }
  }
}

// This is a difference from the docs for the sheets API. Everything comes back as a string, so we cannot
// guess the data type of the values.
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ValueRange {
  pub range: String,
  #[serde(rename = "majorDimension")]
  pub major_dimension: MajorDimension,
  pub values: Vec<Vec<String>>,
}

impl WrapiResult for ValueRange {
  fn parse(_headers: Vec<(String, String)>, body: Vec<u8>) -> Result<Box<ValueRange>, WrapiError> {
    let contents = std::str::from_utf8(&body)?;
    log::debug!("Value Range contents:\n{:#?}", contents);
    let result = serde_json::from_str(contents)?;
    Ok(Box::new(result))
  }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ConditionalFormatRule {
  conditionalformatrule_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum SortOrder {
  #[serde(rename = "SORT_ORDER_UNSPECIFIED")]
  Unspecified,
  #[serde(rename = "ASCENDING")]
  Ascending,
  #[serde(rename = "DESCENDING")]
  Descending,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct SortSpec {
  #[serde(rename = "dimensionIndex")]
  dimension_index: i16,
  #[serde(rename = "sortOrder")]
  sort_order: SortOrder,
  #[serde(rename = "foregroundColor")]
  foreground_color: Color,
  #[serde(rename = "backgroundColor")]
  background_color: Color,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum ConditionType {
  // The default value, do not use.
  #[serde(rename = "CONDITION_TYPE_UNSPECIFIED")]
  ConditionTypeUnspecified,
  //The cell's value must be greater than the condition's value. Supported by data validation , conditional
  // formatting and filters. Requires a single ConditionValue .
  #[serde(rename = "NUMBER_GREATER")]
  NumberGreater,
  // The cell's value must be greater than or equal to the condition's value. Supported by data validation ,
  // conditional formatting and filters. Requires a single ConditionValue .
  #[serde(rename = "NUMBER_GREATER_THAN_EQ")]
  NumberGreaterThanEq,
  // The cell's value must be less than the condition's value. Supported by data validation , conditional
  // formatting and filters. Requires a single ConditionValue .
  #[serde(rename = "NUMBER_LESS")]
  NumberLess,
  // The cell's value must be less than or equal to the condition's value. Supported by data validation ,
  // conditional formatting and filters. Requires a single ConditionValue .
  #[serde(rename = "NUMBER_LESS_THAN_EQ")]
  NumberLessThanEq,
  // The cell's value must be equal to the condition's value. Supported by data validation , conditional formatting
  // and filters. Requires a single ConditionValue .
  #[serde(rename = "NUMBER_EQ")]
  NumberEq,
  // The cell's value must be not equal to the condition's value. Supported by data validation , conditional
  // formatting and filters. Requires a single ConditionValue .
  #[serde(rename = "NUMBER_NOT_EQ")]
  NumberNotEq,
  // The cell's value must be between the two condition values. Supported by data validation , conditional
  // formatting and filters. Requires exactly two ConditionValues .
  #[serde(rename = "NUMBER_BETWEEN")]
  NumberBetween,
  // The cell's value must not be between the two condition values. Supported by data validation , conditional
  // formatting and filters. Requires exactly two ConditionValues .
  #[serde(rename = "NUMBER_NOT_BETWEEN")]
  NumberNotBetween,
  // The cell's value must contain the condition's value. Supported by data validation , conditional formatting
  // and filters. Requires a single ConditionValue .
  #[serde(rename = "TEXT_CONTAINS")]
  TextContains,
  // The cell's value must not contain the condition's value. Supported by data validation , conditional
  // formatting and filters. Requires a single ConditionValue .
  #[serde(rename = "TEXT_NOT_CONTAINS")]
  TextNotContains,
  // The cell's value must start with the condition's value. Supported by conditional formatting and filters.
  // Requires a single ConditionValue .
  #[serde(rename = "TEXT_STARTS_WITH")]
  TextStartsWith,
  // The cell's value must end with the condition's value. Supported by conditional formatting and filters.
  // Requires a single ConditionValue .
  #[serde(rename = "TEXT_ENDS_WITH")]
  TextEndsWith,
  // The cell's value must be exactly the condition's value. Supported by data validation , conditional
  // formatting and filters. Requires a single ConditionValue .
  #[serde(rename = "TEXT_EQ")]
  TextEq,
  // The cell's value must be a valid email address. Supported by data validation. Requires no ConditionValues .
  #[serde(rename = "TEXT_IS_EMAIL")]
  TextIsEmail,
  // The cell's value must be a valid URL. Supported by data validation. Requires no ConditionValues .
  #[serde(rename = "TEXT_IS_URL")]
  TextIsUrl,
  // The cell's value must be the same date as the condition's value. Supported by data validation ,
  // conditional formatting and filters. Requires a single ConditionValue .
  #[serde(rename = "DATE_EQ")]
  DateEq,
  // The cell's value must be before the date of the condition's value. Supported by data validation ,
  // conditional formatting and filters. Requires a single ConditionValue that may be a relative date .
  #[serde(rename = "DATE_BEFORE")]
  DateBefore,
  // The cell's value must be after the date of the condition's value. Supported by data validation , conditional
  // formatting and filters. Requires a single ConditionValue that may be a relative date .
  #[serde(rename = "DATE_AFTER")]
  DateAfter,
  // The cell's value must be on or before the date of the condition's value. Supported by data validation.
  // Requires a single ConditionValue that may be a relative date .
  #[serde(rename = "DATE_ON_OR_BEFORE")]
  DateOnOrBefore,
  // The cell's value must be on or after the date of the condition's value. Supported by data validation.
  // Requires a single ConditionValue that may be a relative date .
  #[serde(rename = "DATE_ON_OR_AFTER")]
  DateOnOrAfter,
  // The cell's value must be between the Dates of the two condition values. Supported by data validation.
  // Requires exactly two ConditionValues .
  #[serde(rename = "DATE_BETWEEN")]
  DateBetween,
  // The cell's value must be outside the Dates of the two condition values. Supported by data validation.
  // Requires exactly two ConditionValues .
  #[serde(rename = "DATE_NOT_BETWEEN")]
  DateNotBetween,
  // The cell's value must be a date. Supported by data validation. Requires no ConditionValues .
  #[serde(rename = "DATE_IS_VALID")]
  DateIsValid,
  // The cell's value must be listed in the grid in condition value's range. Supported by data validation.
  // Requires a single ConditionValue , and the value must be a valid range in A1 notation.
  #[serde(rename = "ONE_OF_RANGE")]
  OneOfRange,
  // The cell's value must be in the list of condition values. Supported by data validation. Supports any number
  // of condition values , one per item in the list. Formulas are not supported in the values.
  #[serde(rename = "ONE_OF_LIST")]
  OneOfList,
  // The cell's value must be empty. Supported by conditional formatting and filters. Requires no ConditionValues .
  #[serde(rename = "BLANK")]
  Blank,
  // The cell's value must not be empty. Supported by conditional formatting and filters. Requires no ConditionValues .
  #[serde(rename = "NOT_BLANK")]
  NotBlank,
  // The condition's formula must evaluate to true. Supported by data validation , conditional formatting and
  // filters. Requires a single ConditionValue .
  #[serde(rename = "CUSTOM_FORMULA")]
  CustomFormula,

  #[serde(rename = "BOOLEAN")]
  Boolean,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum CompDate {
  #[serde(rename = "DATE_BEFORE")]
  DateBefore,
  #[serde(rename = "DATE_AFTER")]
  DateAfter,
  #[serde(rename = "DATE_ON_OR_BEFORE")]
  DateOnOrBefore,
  #[serde(rename = "DATE_ON_OR_AFTER")]
  DateOnOrAfter,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum ConditionValue {
  #[serde(rename = "relativeDate")]
  RelativeDate(CompDate),
  #[serde(rename = "userEnteredValue")]
  UserEnteredValue(String),
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct BooleanCondition {
  #[serde(rename = "type")]
  condition_type: ConditionType,
  values: Vec<ConditionValue>,
}

// TODO: This is a bug. The visibles must be the only set condition
// https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/sheets#filtercriteria
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct FilterCriteria {
  #[serde(default, rename = "hiddenValues")]
  hidden_values: Vec<String>,
  condition: Option<BooleanCondition>,
  visibleBackgroundColor: Option<Color>,
  visibleForegroundColor: Option<Color>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct FilterView {
  #[serde(rename = "filterViewId")]
  filter_view_id: i64,
  title: String,
  range: GridRange,
  #[serde(rename = "namedRangeId")]
  named_range_id: Option<String>,
  #[serde(default, rename = "sortSpecs")]
  sort_specs: Vec<SortSpec>,
  #[serde(default)]
  criteria: HashMap<i64, FilterCriteria>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ProtectedRange {
  protectedrange_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct BasicFilter {
  range: GridRange,
  #[serde(default, rename = "sortSpecs")]
  sort_specs: Vec<SortSpec>,
  #[serde(default)]
  criteria: HashMap<i64, FilterCriteria>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct EmbeddedChart {
  embeddedchart_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct BandedRange {
  bandedrange_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum Dimension {
  #[serde(rename = "DIMENSION_UNSPECIFIED")]
  DimensionUnspecified,
  #[serde(rename = "ROWS")]
  Rows,
  #[serde(rename = "COLUMNS")]
  Columns,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct DimensionRange {
  #[serde(rename = "sheetId")]
  sheet_id: i64,
  dimension: Dimension,
  #[serde(rename = "startIndex")]
  start_index: i32,
  #[serde(rename = "endIndex")]
  end_index: i32,
}
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct DimensionGroup {
  range: DimensionRange,
  depth: i32,
  #[serde(default = "bool_false")]
  collapsed: bool,
}
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct Slicer {
  slicer_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct Sheet {
  properties: SheetProperties,
  #[serde(default)]
  data: Vec<GridData>,
  #[serde(default)]
  merges: Vec<GridRange>,
  #[serde(default, rename = "conditionalFormats")]
  conditional_formats: Vec<ConditionalFormatRule>,
  #[serde(default, rename = "filterViews")]
  filter_views: Vec<FilterView>,
  #[serde(default, rename = "filterViews")]
  protected_ranges: Vec<ProtectedRange>,
  #[serde(rename = "basicFilter")]
  basic_filter: Option<BasicFilter>,
  #[serde(default)]
  charts: Vec<EmbeddedChart>,
  #[serde(default, rename = "bandedRanges")]
  banded_ranges: Vec<BandedRange>,
  #[serde(default, rename = "developerMetadata")]
  developer_metadata: Vec<DeveloperMetadata>,
  #[serde(default, rename = "rowGroups")]
  row_groups: Vec<DimensionGroup>,
  #[serde(default, rename = "columnGroups")]
  column_groups: Vec<DimensionGroup>,
  #[serde(default)]
  slicers: Vec<Slicer>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct NamedRange {
  #[serde(rename = "namedRangeId")]
  named_range_id: String,
  name: String,
  range: GridRange,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct DeveloperMetadata {
  developer_metadata_not_implemented: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct Spreadsheet {
  #[serde(rename = "spreadsheetId")]
  spreadsheet_id: String,
  // properties: SpreadsheetProperties,
  sheets: Vec<Sheet>,
  #[serde(rename = "namedRanges")]
  named_ranges: Vec<NamedRange>,
  #[serde(rename = "spreadsheetUrl")]
  spreadsheet_url: String,
  #[serde(rename = "developerMetadata", default)]
  developer_metadata: Vec<DeveloperMetadata>,
}

impl WrapiResult for Spreadsheet {
  fn parse(_headers: Vec<(String, String)>, body: Vec<u8>) -> Result<Box<Spreadsheet>, WrapiError> {
    // println!("Serde Result:\n{:#?}", std::str::from_utf8(&b ody));
    debug!("Parsing the spreadsheet result");
    let str_body = std::str::from_utf8(&body)?;
    let result: Result<Spreadsheet, serde_json::error::Error> = serde_json::from_str(str_body);
    match result {
      Ok(res) => Ok(Box::new(res)),
      Err(err) => {
        error!("Received an error parsing the GoogleSheet.read response");
        info!("Result: {:#?}", err);
        // debug!("Body: {:#?}", err);
        Err(err)?
      }
    }
  }
}

pub struct Settings {
  auto_write: bool,
}

pub struct OpenRequest {
  sheet_id: String,
}

impl WrapiRequest for OpenRequest {
  fn build_uri(&self, base_url: &str) -> Result<String, WrapiError> {
    let uri = format!("{}{}", base_url, self.sheet_id.clone())
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

pub struct ReadRequest {
  spreadsheet_id: String,
  sheet_name: String,
  range: GridRange,
}

impl WrapiRequest for ReadRequest {
  fn build_uri(&self, base_url: &str) -> Result<String, WrapiError> {
    let uri = url::Url::parse_with_params(
      &format!(
        "{}{}/values/{}",
        base_url,
        self.spreadsheet_id,
        self.range.to_range(self.sheet_name.clone())
      )[..],
      &[("dateTimeRenderOption", "SERIAL_NUMBER")],
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

pub struct SheetDB {
  api: RefCell<wrapi::API>,
  sheet: Box<Spreadsheet>,
  settings: Settings,
}

/// Access a google sheet and keep an API connection open for modifying it
impl SheetDB {
  // pub fn new(&self) -> Result<(), WrapiError> {
  //   Ok(())
  // }

  /// Connect to an existing spreadsheet
  pub fn open(auth: wrapi::AuthMethod, sheet_id: String) -> Result<SheetDB, WrapiError> {
    info!("Opening spreadsheet with ID: {}", sheet_id.clone());
    let api = wrapi::API::new(auth.clone())
      .add_endpoint(
        "open".to_string(),
        wrapi::Endpoint {
          base_url: "https://sheets.googleapis.com/v4/spreadsheets/",
          auth_method: auth.clone(),
          request_method: wrapi::RequestMethod::GET,
          scopes: vec!["https://www.googleapis.com/auth/drive"],
          request_mime_type: wrapi::MimeType::Json,
          response_mime_type: wrapi::MimeType::Json,
        },
      )
      .add_endpoint(
        "read".to_string(),
        wrapi::Endpoint {
          base_url: "https://sheets.googleapis.com/v4/spreadsheets/",
          auth_method: auth.clone(),
          request_method: wrapi::RequestMethod::GET,
          scopes: vec!["https://www.googleapis.com/auth/drive"],
          request_mime_type: wrapi::MimeType::Json,
          response_mime_type: wrapi::MimeType::Json,
        },
      );

    let req = OpenRequest { sheet_id: sheet_id };
    debug!("About to query the sheet");
    let sheet = api.call("open", req);

    Ok(SheetDB {
      api: RefCell::new(api),
      sheet: sheet.unwrap(),
      settings: Settings { auto_write: false },
    })
  }

  // list sheets
  pub fn list_sheets(&self) -> Result<Vec<String>, WrapiError> {
    let list = self
      .sheet
      .sheets
      .iter()
      .map(|item| item.properties.title.clone())
      .collect();
    Ok(list)
  }

  // pub fn get_range(&self, range: GridRange)

  /// column headers. This should likely be in Subpar, since this suggests a specific format for the worksheet
  pub fn get_headers(&self, sheet_name: String) -> Result<Vec<String>, WrapiError> {
    debug!("Reading headers the sheet_name:\n{:#?}", sheet_name);
    let data = self.sheet.sheets.iter().find_map(|sheet| {
      match sheet.properties.title.clone() == sheet_name {
        true => {
          let req = ReadRequest {
            spreadsheet_id: self.sheet.spreadsheet_id.clone(),
            sheet_name: sheet_name.clone(),
            range: GridRange {
              sheet_id: sheet.properties.sheet_id,
              start_row_index: 1,
              start_column_index: 1,
              end_row_index: 1,
              end_column_index: sheet.properties.grid_properties.column_count,
            },
          };

          let data: Result<Box<ValueRange>, WrapiError> = self.api.borrow_mut().call("read", req);
          Some(data)
        }
        false => None,
      }
    });
    debug!("Read Header data:\n{:#?}", data);
    match data {
      Some(Ok(x)) => Ok(match x {
        _ => Err(WrapiError::General(
          "Data Conversion of {} does not make sense when getting data from Sheets".to_string(),
        ))?,
      }),
      Some(Err(err)) => Err(err),
      None => Err(WrapiError::General(format!(
        "Could not find a sheet named '{}' to get_headers from",
        sheet_name
      ))),
    }
  }

  pub fn get_sheet_properties(&self, sheet_name: String) -> Result<SheetProperties, WrapiError> {
    debug!("Reading headers the sheet_name:\n{:#?}", sheet_name);
    let data = self.sheet.sheets.iter().find_map(|sheet| {
      match sheet.properties.title.clone() == sheet_name {
        true => Some(sheet.properties.clone()),
        false => None,
      }
    });
    match data {
      Some(x) => Ok(x),
      None => Err(WrapiError::General(format!(
        "Sheet named '{}' not found",
        sheet_name
      ))),
    }
  }

  pub fn get_sheet(&self, sheet_name: String) -> Result<Box<ValueRange>, WrapiError> {
    let props = self.get_sheet_properties(sheet_name.clone())?;
    let req = ReadRequest {
      spreadsheet_id: self.sheet.spreadsheet_id.clone(),
      sheet_name: sheet_name.clone(),
      range: GridRange {
        sheet_id: props.sheet_id,
        start_row_index: 1,
        start_column_index: 1,
        end_row_index: 2,
        // end_row_index: props.grid_properties.row_count,
        end_column_index: props.grid_properties.column_count,
      },
    };

    self.api.borrow_mut().call("read", req)
  }

  // get row

  // update cell
}
