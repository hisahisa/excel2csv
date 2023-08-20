use chrono::{Duration, NaiveDateTime};

fn excel_date_to_datetime(f_val: &str, excel_base_date: NaiveDateTime) -> String {
    // エクセルの数値を日数と秒に分割
    let days: i64 = f_val.split('.').next().unwrap().parse().unwrap();
    let seconds: i64 = ((f_val.parse::<f64>().unwrap() - days as f64) * 86400.5) as i64;

    // エクセルの基準日に指定された秒数を加算
    let result = excel_base_date +
        Duration::days(days) +
        Duration::seconds(seconds);

    // 時刻の部分を取得し、文字列としてフォーマット
    let time_format = result.format(if f_val.contains('.') {
        "%Y-%m-%d %H:%M:%S"
    } else {
        "%Y-%m-%d"
    }).to_string();
    time_format
}

fn main() {
    // エクセルの基準日 (1900年1月1日)
    let excel_base_date = NaiveDateTime::new(
        chrono::NaiveDate::from_ymd(1970, 12, 30),
        chrono::NaiveTime::from_hms(0, 0, 0));

    // エクセルの数値
    let f_val = "44995.417500000003";

    // 変換して表示
    let result = excel_date_to_datetime(f_val, excel_base_date);
    println!("{}", result);
}
