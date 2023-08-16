use std::borrow::Cow;
use chrono::{Duration, NaiveDateTime};
use quick_xml::name::QName;


#[derive( Debug, Clone)]
pub(crate) struct StructCsv {
    value: String,
    attr: u8,
//    attr_v: Cow<'a, [u8]>,
}

impl StructCsv {
    //pub(crate) fn new (attr: QName<'a>, attr_v: Cow<'a, [u8]>) -> StructCsv<'a> {
    pub(crate) fn new (attr: u8) -> StructCsv {
        StructCsv{
            value: "".to_string(),
            attr,
//            attr_v
        }
    }
    // pub(crate) fn setAttr(&mut self, attr: String, attr_v: String) {
    //     self.attr = attr;
    //     self.attr_v = attr_v;
    // }
    pub(crate) fn setValue(&mut self, val: String) {
        self.value = val;
    }
}