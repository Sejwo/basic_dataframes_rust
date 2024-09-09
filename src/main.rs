// Import necessary modules
use std::fmt;
use calamine::{ Reader, open_workbook, Xlsx, DataType, Data, ExcelDateTime };

// Enums

/// Custom error type for date parsing.
#[derive(Debug)]
enum DateParseError {
    UnsupportedFormat,
    InvalidDay,
    InvalidMonth,
    InvalidYear,
    InvalidSerialNumber, // Added for better error handling
    InvalidDate, // Added for date conversion errors
}

/// Enum for different types a cell can have.
#[derive(Debug)]
enum CellValues {
    Int(i32),
    Float(f64),
    Text(String),
    Date(Date),
}

//traits
// Implement the conversion for different types
impl From<i32> for CellValues {
    fn from(value: i32) -> Self {
        CellValues::Int(value)
    }
}

impl From<f64> for CellValues {
    fn from(value: f64) -> Self {
        CellValues::Float(value)
    }
}

impl From<String> for CellValues {
    fn from(value: String) -> Self {
        CellValues::Text(value)
    }
}

// Also allow &str to be converted to Text as owned
impl From<&str> for CellValues {
    fn from(value: &str) -> Self {
        CellValues::Text(value.to_string())
    }
}
// Structs

///Struct to hold CellValues
#[derive(Debug)]
struct Cell {
    value: Option<CellValues>,
}

/// Struct to represent a Date.
#[derive(Debug)]
struct Date {
    year: u32,
    month: u8,
    day: u8,
}

/// Struct for DataFrame which uses the Cell enum.
#[derive(Debug)]
struct DataFrame {
    data: Vec<Vec<Cell>>,
}

// Implementations
impl DataFrame {
    fn handle_dates(data: &Data) -> Option<Date> {
        match data {
            Data::DateTime(excel_date_time) => {
                // Check if value can be converted to f64 and then to u32
                let serial_number = excel_date_time.as_f64() as u32;
                let date = Date::from_excel_datetype(serial_number).ok();
                println!("Date extracted: {:?}", date);
                date
            }
            _ => {
                println!("Not a DateTime: {:?}", data);
                None
            }
        }
    }

    pub fn new(data: Vec<Vec<Cell>>) -> Self {
        DataFrame { data }
    }

    pub fn read_from_xlsx(
        &mut self,
        path: &str,
        provided_sheet_name: Option<&str>,
        provided_with_headers: Option<bool>
    ) {
        let mut data_for_dataframe: Vec<Vec<Cell>> = vec![];
        let sheet_name = provided_sheet_name.unwrap_or("Sheet1");
        let mut workbook: Xlsx<_> = open_workbook(path).expect("failed to open file");

        if let Ok(range) = workbook.worksheet_range(sheet_name) {
            for rows in range.rows() {
                let mut temp_row: Vec<Cell> = vec![];
                for individual_cell in rows {
                    // Convert cell data into CellValues
                    let temp_cell = into_cell_value(individual_cell);

                    if let Some(cell_value) = temp_cell {
                        temp_row.push(Cell { value: Some(cell_value) });
                    } else {
                        // Try to handle as date if not already handled
                        if let Some(date) = Self::handle_dates(individual_cell) {
                            temp_row.push(Cell { value: Some(CellValues::Date(date)) });
                        } else {
                            // Handle cases where `individual_cell` is `None`
                            temp_row.push(Cell { value: None });
                        }
                    }
                }
                // Check the contents of temp_row for debugging
                data_for_dataframe.push(temp_row);
            }
        } else {
            println!("Failed to read worksheet range");
        }
        self.data = data_for_dataframe;
    }
}

impl Date {
    /// Parses date components based on the given format.
    fn is_leap_year(year: u32) -> bool {
        (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
    }
    fn days_in_year(year: u16) -> u32 {
        if Self::is_leap_year(year.into()) { 366 } else { 365 }
    }
    fn days_in_month(year: u16, month: u8) -> u8 {
        match month {
            1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
            4 | 6 | 9 | 11 => 30,
            2 => if Self::is_leap_year(year.into()) { 29 } else { 28 }
            _ => 0,
        }
    }
    pub fn from_excel_datetype(serial: u32) -> Result<Self, DateParseError> {
        if serial < 1 {
            return Err(DateParseError::InvalidSerialNumber);
        }

        let base_year = 1900;
        let corrected_serial = if serial > 60 { serial - 1 } else { serial };
        let days_since_base = corrected_serial - 1;

        let mut year = base_year;
        let mut days_remaining: u32 = days_since_base;

        while days_remaining >= Self::days_in_year(year) {
            days_remaining -= Self::days_in_year(year);
            year += 1;
        }

        let mut month = 1;
        while days_remaining >= Self::days_in_month(year, month).into() {
            days_remaining -= Self::days_in_month(year, month) as u32;
            month += 1;
        }

        let day = (days_remaining as u8) + 1;
        if day > Self::days_in_month(year, month) {
            return Err(DateParseError::InvalidDate);
        }

        Ok(Date { year: year.into(), month: month.into(), day: day.into() })
    }
    pub fn from_numbers<T>(
        frag1: T,
        frag2: T,
        frag3: T,
        format: &str
    ) -> Result<Self, DateParseError>
        where T: Into<u32>
    {
        let months_with_31_days: [u8; 7] = [1, 3, 5, 7, 8, 10, 12];
        let (year, month, day) = match format {
            "YYYY/MM/DD" => (frag1.into(), frag2.into() as u8, frag3.into() as u8),
            "DD/MM/YYYY" => (frag3.into(), frag2.into() as u8, frag1.into() as u8),
            "MM/DD/YYYY" => (frag3.into(), frag1.into() as u8, frag2.into() as u8),
            _ => {
                return Err(DateParseError::UnsupportedFormat);
            }
        };

        // validation logic
        if day >= 32 {
            return Err(DateParseError::InvalidDay);
        }
        if !months_with_31_days.contains(&month) && !(day <= 31) {
            return Err(DateParseError::InvalidDay);
        }
        if month == (2 as u8) {
            if Self::is_leap_year(year) {
                if day > 29 {
                    return Err(DateParseError::InvalidDay);
                }
            } else {
                if day > 28 {
                    return Err(DateParseError::InvalidDay);
                }
            }
        }

        if !(1..=12).contains(&month) {
            return Err(DateParseError::InvalidMonth);
        }
        if year == 0 {
            return Err(DateParseError::InvalidYear);
        }

        Ok(Date { year, month, day })
    }
}

impl fmt::Display for DateParseError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            DateParseError::UnsupportedFormat => write!(f, "Unsupported date format"),
            DateParseError::InvalidDay => write!(f, "Invalid day"),
            DateParseError::InvalidMonth => write!(f, "Invalid month"),
            DateParseError::InvalidYear => write!(f, "Invalid year"),
            &DateParseError::InvalidSerialNumber | &DateParseError::InvalidDate => todo!(),
        }
    }
}

///utils

///assert what is the type of the cell in worksheet
fn into_cell_value(data: &dyn DataType) -> Option<CellValues> {
    if let Some(val) = data.get_int() {
        if val >= (i32::MIN as i64) && val <= (i32::MAX as i64) {
            return Some(CellValues::Int(val as i32));
        }
    } else if let Some(val) = data.get_float() {
        return Some(CellValues::Float(val));
    } else if let Some(val) = data.get_string() {
        return Some(CellValues::Text(val.to_string()));
    } else if let Some(val) = data.get_bool() {
        return Some(CellValues::Text(val.to_string()));
    } else if data.is_empty() {
        return None;
    } else if let Some(err) = data.get_error() {
        println!("Error in data: {:?}", err);
        return None;
    }

    None
}

// Main function for testing and debugging
fn main() {
    let date: Date = Date::from_numbers::<u32>(4, 2, 2000, "DD/MM/YYYY").expect(
        "Error when parsing"
    );
    println!("{:?}", date);
    let mut dftest = DataFrame::new(vec![vec![]]);
    println!("{:?}", dftest);
    dftest.read_from_xlsx("data\\test.xlsx", Some("Sheet1"), Some(true));
    println!("{:#?}", dftest)
}
