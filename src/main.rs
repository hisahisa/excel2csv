mod structual;

use std::fs::File;
use std::io::Read;
use zip::read::{ZipArchive};
use quick_xml::Reader;
use quick_xml::events::Event;
use chrono::NaiveDateTime;
use crate::structual::StructCsv;

fn read_xml_inside_zip(mut zip: ZipArchive<&File>, folder: &str, sheet_name: &str)
                       -> Vec<Vec<StructCsv>> {

    // 指定されたシート名検索
    let folder2 = "worksheets";
    let inner_zip_entry_path = format!("{}/{}/{}.xml", folder, folder2, sheet_name);
    let mut inner_file = zip.by_name(&inner_zip_entry_path).unwrap();

    // ファイルの内容を読み込む
    let mut content = Vec::new();
    inner_file.read_to_end(&mut content).unwrap();

    // 読み取った内容をXMLとして解析して表示
    let mut xml_reader = Reader::from_reader(&content[..]);
    let mut buffer = Vec::new();
    let mut c_list: Vec<Vec<StructCsv>> = Vec::new();
    let mut row: Vec<StructCsv> = Vec::new();
    let mut is_v = false;
    let mut s: Option<StructCsv> = None;
    let target_attr = vec![115u8, 116u8];  // b"s", b"t"
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
                        c_list.push(row);
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

fn string_resolve_vec(mut zip: ZipArchive<&File>, folder_name: &str) ->  Vec<String> {
    let inner_zip_str_file = "sharedStrings.xml";
    let inner_zip_str_path = format!("{}/{}", folder_name, inner_zip_str_file);
    let mut zip_file = zip.by_name(&inner_zip_str_path).unwrap();

    // ファイルの内容を読み込む
    let mut content = Vec::new();
    zip_file.read_to_end(&mut content).unwrap();

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
    // 元のZIPファイルのパスと読み込むフォルダ名とファイル名を指定
    let zip_path = "tenpo_shohin_pattern_x.xlsx";
    let zip_path = "tenpo_shohin_pattern3.xlsx";
    let sheet_name = "sheet1";

    // excelフォルダ
    let folder = "xl";
    // zipファイルを開く
    let file = File::open(zip_path).unwrap();
    let zip = ZipArchive::new(&file).unwrap();
    let zip_str_sol = ZipArchive::new(&file).unwrap();

    // 名前解決リスト作成
    let name_resolve: Vec<String> = string_resolve_vec(zip_str_sol, folder);
    println!("Start: {:?}", name_resolve);

    let xml_contents = read_xml_inside_zip(zip, folder, sheet_name);
    let navi = create_navi();
    let mut csv_str: Vec<String> = Vec::new();
    for i in xml_contents {
        let res = i.into_iter().map(|a| {
            a.clone().get_value(&navi, &name_resolve)
        }).collect::<Vec<String>>().join(",");
        csv_str.push(res);
    }
    println!("ok");
    println!("csv_str: {:?}", csv_str.join("\n"));
}
