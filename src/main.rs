use regex::Regex;
use docx::document::Paragraph;
use docx::Docx;
use calamine::{Reader, open_workbook, Xlsx, Range, CellType};

enum ColumnFilter {
    Equals,
    Greater(i32),
    Less(i32),
    GreaterEq(i32),
    LessEq(i32),
}

macro_rules! read {
    ($out:ident as $type:ty) => {
        let mut inner = String::new();
        std::io::stdin().read_line(&mut inner).expect("not as string");
        let $out = inner.trim().parse::<$type>().expect("not parsable");
    };
}

macro_rules! read_str {
    ($out:ident) => {
        let mut inner = String::new();
        std::io::stdin().read_line(&mut inner).expect("not a string");
        let $out = inner.trim();
    };
}

macro_rules! read_vec {
    ($out:ident as $type:ty) => {
        let mut inner = String::new();
        std::io::stdin().read_line(&mut inner).unwrap();
        let $out = inner
            .trim()
            .split_whitespace()
            .map(|s| s.parse::<$type>().unwrap())
            .collect::<Vec<$type>>();
    };
}

fn parse_side(a: &str) -> (u32, u32) {
    let mut vec_x: Vec<u32> = Vec::new();
    let mut vec_y: Vec<u32> = Vec::new();
    for byte in a.as_bytes() {
        match byte - b'0' {
            0..=9 => vec_y.push((byte - b'0').into()),
            _ => vec_x.push((byte - b'A').into()),
        }
    }
    let x: u32 = vec_x.iter().rev().enumerate().map(|(index, value)| 26u32.pow(index as u32) * (value + 1)).sum();
    let y: u32 = vec_y.iter().rev().enumerate().map(|(index, value)| 10u32.pow(index as u32) * value).sum();
    (x - 1, y - 1)
}

fn to_coords(a: String) -> (u32, u32, u32, u32) {
    let sides: Vec<&str> = a.split(":").collect();
    let lhs = sides[0];
    let rhs = sides[1];
    let (x1, y1) = parse_side(lhs);
    let (x2, y2) = parse_side(rhs);
    (x1, y1, x2, y2)
}

fn print_header<T: CellType + std::fmt::Display>(x_start: u32, x_end: u32, y: u32, range: &Range<T>) {
    for x in x_start..=x_end {
        let value = range.get_value((y, x));
        match value {
            None => print!("-\t\t"),
            Some(x) => print!("{}\t\t", x.to_string())
        }
    }
    println!();
}

fn main() {
    println!("Введите путь к входному файлу: ");
    read_str!(input_path);
    let mut workbook: Xlsx<_> = open_workbook(&input_path).expect("Cannot open file");
    println!("Введите область значений (MX:NY): ");
    read_str!(str_range);
    let (x1, mut y1, x2, y2) = to_coords(str_range.to_string());
    let parse_i32 = { 
        let numbers_re = Regex::new("[+-]?([0-9]*[.])?[0-9]+").unwrap();
        move |x: &str| numbers_re.find(x).unwrap().as_str().parse::<i32>().unwrap()
    };
    // println!("{} {} {} {}", x1, y1, x2, y2);
    if let Some(Ok(range)) = workbook.worksheet_range(&workbook.sheet_names().to_owned()[0]) {
        println!("Список колонок: ");
        (1..=(x2 - x1 + 1)).for_each(|x| print!("{}\t\t", x));
        println!();
        print_header(x1, x2, y1, &range);
        y1 += 1;
        println!("Выберите номера колонок (через пробел): ");
        read_vec!(chosen_columns as u32);
        let mut expressions: Vec<Regex> = Vec::new();
        for i in &chosen_columns {
            println!("Регулярное выражение для колонки {}:", i);
            read_str!(reg);
            expressions.push(Regex::new(reg).unwrap());
        }
        let mut filters: Vec<ColumnFilter> = Vec::new();
        for i in &chosen_columns {
            println!("Введите фильтр ('=', '>' | '<' | '<=' | '>=' x для {}:", i);
            read_str!(filter);
            if filter.starts_with("=") {
                filters.push(ColumnFilter::Equals);
            } else if filter.starts_with(">=") {
                let number = parse_i32(filter);
                filters.push(ColumnFilter::GreaterEq(number));
            } else if filter.starts_with("<=") {
                let number = parse_i32(filter);
                filters.push(ColumnFilter::LessEq(number));
            } else if filter.starts_with(">") {
                let number = parse_i32(filter);
                filters.push(ColumnFilter::Greater(number));
            } else if filter.starts_with("<") {
                let number = parse_i32(filter);
                filters.push(ColumnFilter::Less(number));
            } else {
                panic!("Неправильный фильтр");
            }
        }
        println!("Выберите колонку, которая будет выводиться в файл");
        read!(out_column as u32);
        let mut to_word: Vec<String> = Vec::new();
        for y in y1..=y2 {
            if chosen_columns.iter().enumerate().all(|(index, x)| {
                let reg = &expressions[index];
                let filter = &filters[index];
                if let Some(cell) = range.get_value((y, x1 + x - 1)) {
                    let cell = cell.to_string();
                    // println!("{}", cell);
                    if let Some(mat) = reg.find(&cell) {
                        // println!("{}", mat.as_str());
                        use ColumnFilter::*;
                        match filter {
                            Equals => true,
                            Greater(val) => parse_i32(mat.as_str()) > *val,
                            Less(val) => parse_i32(mat.as_str()) < *val,
                            GreaterEq(val) => parse_i32(mat.as_str()) >= *val,
                            LessEq(val) => parse_i32(mat.as_str()) <= *val,
                        }
                    } else {
                        false
                    }
                } else {
                    false
                }
            })
            {
                match range.get_value((y, x1 + out_column - 1)) {
                    Some(x) => to_word.push(x.to_string()),
                    None => to_word.push(" - ".into())
                }
            }
        }
        println!("{}", to_word.join(", "));
        println!("Введите путь к выходному файлу: ");
        read_str!(output_path);
        let mut docx = Docx::default();
        let head = Paragraph::default().push_text(
            (0..chosen_columns.len()).map(|index| {
                let reg = &expressions[index];
                let filter = &filters[index];

                use ColumnFilter::*;

                match filter {
                    Equals => reg.as_str().to_string(),
                    Greater(val) => format!("более {}", val),
                    Less(val) => format!("менее {}", val),
                    GreaterEq(val) => format!("от {}", val),
                    LessEq(val) => format!("до {}", val),
                }
            }).collect::<Vec<String>>().join(", ")
        ).push_text(": ");
        docx.document.push(head);
        let body = Paragraph::default().push_text(
            to_word.join(", ")
        );
        docx.document.push(body);
        docx.write_file(&output_path).unwrap();
    }
}
