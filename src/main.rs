use log::{debug, error, trace};
use regex::Regex;
use std::{
    collections::HashMap,
    fmt, fs,
    io::{self, Read},
    path::Path,
};

use clap::Parser;
use env_logger;
use once_cell::sync::Lazy;
use serde::Serialize;
use zip::{self, read::ZipFile};

static RE_TEXT: Lazy<Regex> = Lazy::new(|| Regex::new(r"<a:t>([\s\S]+?)</a:t>").unwrap());
static RE_PAGE_NO: Lazy<Regex> = Lazy::new(|| Regex::new(r"(slide|slideMaster)(\d+).xml").unwrap());

#[derive(Parser)]
#[command(version, about, long_about = None)]
struct Args {
    /// Input file
    #[arg(short, long)]
    input_file: String,

    /// Output directory
    #[arg(short, long, default_value_t = String::from("./output"))]
    output_dir: String,
}

const DIR_MEDIA: &str = "ppt/media";
const DIR_SLIDES_RELS: &str = "ppt/slides/_rels";
const MASTER_RELS_DIR: &str = "ppt/slideMasters/_rels";
const DIR_SLIDES: &str = "ppt/slides";
const INDEX_FILE: &str = "index.json";
const ATTR_REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

#[derive(Debug, Serialize)]
struct SingleRes {
    page_no: u32,
    slide_master: bool,
    images: Vec<String>,
    texts: Vec<String>,
}

#[derive(Debug, Serialize)]
struct PageRes {
    slides: HashMap<u32, SingleRes>,
    masters: HashMap<u32, SingleRes>,
}

#[derive(Debug, Serialize)]
struct Res<'a> {
    doc_title: &'a str,
    pages: PageRes,
}

fn main() {
    env_logger::init();
    let args = Args::parse();

    let mut res = Res {
        doc_title: Path::new(&args.input_file)
            .file_name()
            .unwrap()
            .to_str()
            .unwrap(),
        pages: PageRes {
            slides: HashMap::new(),
            masters: HashMap::new(),
        },
    };

    let archivef = fs::File::open(Path::new(&args.input_file)).expect("failed to open input file");
    let freader = std::io::BufReader::new(archivef);
    let mut archive = zip::ZipArchive::new(freader).expect("failed to open archive");

    for i in 0..archive.len() {
        let mut file: ZipFile = archive.by_index(i).unwrap();
        if file.is_dir() {
            continue;
        }

        let fname = file.name().to_owned();
        if fname.starts_with(DIR_MEDIA) {
            match export_media(&Path::new(&args.output_dir), &mut file) {
                Ok(()) => {
                    trace!("exported media {}", fname)
                }
                Err(e) => {
                    error!("failed to export media: {}, error: {}", fname, e)
                }
            };
        } else if fname.starts_with(DIR_SLIDES_RELS) {
            match rels(file) {
                Ok((page_no, rels)) => {
                    trace!("got page {:?}, rels: {:?}", page_no, rels);
                    let page_res = res
                        .pages
                        .slides
                        .entry(page_no)
                        .or_insert_with(|| SingleRes {
                            page_no,
                            slide_master: false,
                            images: Vec::new(),
                            texts: Vec::new(),
                        });
                    page_res.images = rels.values().cloned().collect();
                }
                Err(e) => {
                    error!("failed to get rels, error: {}", e)
                }
            }
        } else if fname.starts_with(DIR_SLIDES) {
            trace!("file {:?} is slide rels", fname);
            let page_res = slide(file);
            if page_res.is_err() {
                error!("failed to get slide, error: {}", page_res.unwrap_err());
                continue;
            }
            let page_res = page_res.unwrap();
            trace!(
                "got page {:?}, texts: {:?}",
                page_res.page_no,
                page_res.texts
            );
            let single_res =
                res.pages
                    .slides
                    .entry(page_res.page_no)
                    .or_insert_with(|| SingleRes {
                        page_no: (&page_res).page_no,
                        slide_master: false,
                        images: Vec::new(),
                        texts: (&page_res).texts.clone(),
                    });
            single_res.texts = page_res.texts.clone();
        } else if fname.starts_with(MASTER_RELS_DIR) {
            match rels(file) {
                Ok((page_no, rels)) => {
                    trace!("got page {:?}, rels: {:?}", page_no, rels);
                    let page_res = res
                        .pages
                        .masters
                        .entry(page_no)
                        .or_insert_with(|| SingleRes {
                            page_no,
                            slide_master: true,
                            images: Vec::new(),
                            texts: Vec::new(),
                        });
                    page_res.images = rels.values().cloned().collect();
                }
                Err(e) => {
                    error!("failed to get rels, error: {}", e)
                }
            }
        }
    }
    debug!("res: {:?}", res);
    let j = serde_json::to_string_pretty(&res).unwrap();
    // write j to {output_dir}/{INDEX_FILE}
    fs::write(Path::new(&args.output_dir).join(INDEX_FILE), j).unwrap();
}

#[derive(Debug)]
enum ExportMediaError {
    Io(std::io::Error),
    Parse(xmltree::ParseError, String),
    Custom(String),
}

impl fmt::Display for ExportMediaError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ExportMediaError::Io(err) => write!(f, "IO error: {}", err),
            ExportMediaError::Parse(err, fname) => write!(f, "Parse error: {} in {}", err, fname),
            ExportMediaError::Custom(err) => write!(f, "Custom error: {}", err),
        }
    }
}

impl std::error::Error for ExportMediaError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            ExportMediaError::Io(err) => Some(err),
            ExportMediaError::Parse(err, _) => Some(err),
            ExportMediaError::Custom(_) => None,
        }
    }
}

impl From<std::io::Error> for ExportMediaError {
    fn from(error: std::io::Error) -> Self {
        ExportMediaError::Io(error)
    }
}

impl From<String> for ExportMediaError {
    fn from(error: String) -> Self {
        ExportMediaError::Custom(error)
    }
}

fn export_media(output: &Path, f: &mut ZipFile) -> Result<(), ExportMediaError> {
    // get the filename from f
    let filename = Path::new(f.name()).file_name().unwrap();
    let outfilename = output.join(filename);
    trace!("out filename: {:?}", outfilename);
    // write contents of f to outfilename
    let mut outfile = fs::File::create(outfilename).map_err(|e| e.to_string())?;

    match io::copy(f, &mut outfile) {
        Ok(_) => Ok(()),
        Err(e) => Err(e.into()),
    }
}

fn slide(mut f: ZipFile) -> Result<SingleRes, String> {
    let mut res = SingleRes {
        page_no: 0,
        slide_master: false,
        images: Vec::new(),
        texts: Vec::new(),
    };
    let mut content: String = String::new();
    f.read_to_string(&mut content).map_err(|e| e.to_string())?;
    for cap in RE_TEXT.captures_iter(&content) {
        if let Some(text) = cap.get(1) {
            res.texts.push(text.as_str().to_owned());
        }
    }
    let fname = f.name();
    res.page_no = page_no(fname)?;
    debug!("page res: {:?}", res);
    Ok(res)
}

fn rels(f: zip::read::ZipFile) -> Result<(u32, HashMap<String, String>), ExportMediaError> {
    let fname = f.name().to_owned();
    let el = xmltree::Element::parse(f).map_err(|e| ExportMediaError::Parse(e, fname.clone()))?;
    let image_rel_nodes = el.children.into_iter().filter(|node: &xmltree::XMLNode| {
        let el = node.as_element().unwrap();
        el.name == "Relationship"
            && el.attributes.get("Type") == Some(&ATTR_REL_TYPE_IMAGE.to_string())
    });
    let mut res = HashMap::new();
    for image_rel_node in image_rel_nodes {
        let image_rel_el = image_rel_node.as_element().unwrap();
        let rel_image_path = image_rel_el.attributes.get("Target").unwrap();
        res.insert(
            image_rel_el.attributes.get("Id").unwrap().to_owned(),
            Path::new(rel_image_path)
                .file_name()
                .unwrap()
                .to_string_lossy()
                .into_owned(),
        );
    }
    let page_no = page_no(&fname)?;
    Ok((page_no, res))
}

// get page no from filename
fn page_no(fname: &str) -> Result<u32, String> {
    if let Some(matched) = RE_PAGE_NO.captures(fname) {
        if let Some(page_no) = matched.get(2) {
            Ok(page_no.as_str().parse::<u32>().unwrap())
        } else {
            Err("Can't find valid page no".into())
        }
    } else {
        Err("Can't find valid page no".into())
    }
}
