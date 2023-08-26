mod structual;

use std::fs::File;
use std::io::Read;
use zip::read::{ZipArchive};
use quick_xml::Reader;
use quick_xml::events::Event;
use chrono::NaiveDateTime;
use crate::structual::StructCsv;

fn read_xml(content: Vec<u8>, name_resolve: &Vec<String>) -> Vec<String> {

    // 読み取った内容をXMLとして解析して表示
    let mut xml_reader = Reader::from_reader(&content[..]);
    let mut buffer = Vec::new();
    let mut c_list: Vec<String> = Vec::new();
    let mut row: Vec<StructCsv> = Vec::new();
    let mut is_v = false;
    let mut s: Option<StructCsv> = None;
    let target_attr = vec![115u8, 116u8];  // b"s", b"t"
    let navi = create_navi();
    loop {
        match xml_reader.read_event_into(&mut buffer) {
            Ok(Event::Start(e)) => {
                match e.name().as_ref() {
                    b"c" => {
                        for i in e.attributes() {
                            match i {
                                Ok(x) => {
                                    let a = if target_attr.
                                        contains(&x.key.into_inner()[0]) {
                                        x.key.into_inner()[0].clone()
                                    } else { 0u8 };
                                    s = Some(StructCsv::new(a));
                                }
                                Err(_) => {}
                            }
                        }
                    },
                    b"v" => is_v = true,
                    _ => {},
                }
            }
            Ok(Event::End(e)) => {
                match e.name().as_ref() {
                    b"row" => {
                        let i = row.into_iter().map(|a| {
                            a.clone().get_value(&navi, &name_resolve)
                        }).collect::<Vec<String>>().join(",");
                        c_list.push(i);
                        row = Vec::new();
                    },
                    _ => {},
                }
            }
            Ok(Event::Text(e)) => {
                if is_v {
                    match s {
                        Some(ref mut v) => {
                            let val = e.unescape().unwrap().into_owned();
                            v.set_value(val);
                            row.push(v.clone());
                            s = None;
                        }
                        None => {}
                    }
                    is_v = false;
                }
            }
            Ok(Event::Eof) => {
                break;
            }
            Err(e) => {
                eprintln!("Error: {}", e);
                break;
            }
            _ => {}
        }
    }
    c_list
}

fn str_resolve(content: Vec<u8>) ->  Vec<String> {

    // 読み取った内容をXMLとして解析して表示
    let mut xml_reader = Reader::from_reader(&content[..]);
    let mut buffer = Vec::new();
    let mut name_resolve: Vec<String> = Vec::new();
    let mut is_text = false;
    let mut no_text_ = false;
    loop {
        match xml_reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref e)) => {
                match e.name().as_ref() {
                    b"t" => is_text = true,
                    b"rPh" => no_text_ = true,
                    _ => no_text_ = false,
                }
            }
            Ok(Event::Text(e)) => {
                if &is_text & !&no_text_ {
                    let val = e.unescape().unwrap().into_owned();
                    name_resolve.push(val);
                    is_text = false;
                    no_text_ = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                eprintln!("Error: {}", e);
                break;
            }
            _ => {}
        }
    }
    name_resolve
}

fn create_navi() -> NaiveDateTime {
    NaiveDateTime::new(
        chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap(),
        chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap())
}

fn main() {
    // excelファイルの指定
    let zip_path = "tenpo_shohin_pattern_x.xlsx";
    let zip_path = "tenpo_shohin_pattern3.xlsx";

    // excelフォルダ
    let folder = "xl";
    // エクセル(zip)ファイルを開く
    let file = File::open(zip_path).unwrap();
    let mut zip = ZipArchive::new(&file).unwrap();
    let mut zip_str_sol = ZipArchive::new(&file).unwrap();

    // 名前解決リスト作成
    let inner_zip_str_file = "sharedStrings.xml";
    let inner_zip_str_path = format!("{}/{}", folder, inner_zip_str_file);
    let mut str_file = zip_str_sol.by_name(&inner_zip_str_path).unwrap();
    let mut str_content = Vec::new();
    str_file.read_to_end(&mut str_content).unwrap();
    let name_resolve: Vec<String> = str_resolve(str_content);
    println!("Start: {:?}", name_resolve);

    // excelコンテンツを読み込む
    let folder2 = "worksheets";
    let sheet_name = "sheet1";
    let inner_zip_entry_path = format!("{}/{}/{}.xml", folder, folder2, sheet_name);
    let mut inner_file = zip.by_name(&inner_zip_entry_path).unwrap();
    let mut content = Vec::new();
    inner_file.read_to_end(&mut content).unwrap();
    let xml_contents = read_xml(content, &name_resolve);
    let csv_str = xml_contents.join("\n");

    println!("ok");
    println!("csv_str: {}", csv_str);
}
