use clap::Parser;
use rust_xlsxwriter::Workbook;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, HashMap};
use std::error::Error;
use std::path::PathBuf;
use std::{fs, io};
use walkdir::WalkDir;

use crate::top_percentile::{F64Ord, TopPercentile};

pub mod top_percentile;

fn parse_xlsx_path(s: &str) -> Result<PathBuf, String> {
    let path = PathBuf::from(s);
    match path.extension().and_then(|e| e.to_str()) {
        Some(ext) if ext.eq_ignore_ascii_case("xlsx") => Ok(path),
        _ => Err(String::from("output must have .xlsx extension")),
    }
}

#[derive(Parser, Debug)]
#[command(
    name = "dp_output_analysis",
    version,
    about = "Analyze .bin representations in a directory",
    arg_required_else_help = true
)]
struct Args {
    #[arg(value_name = "DIR", required = true)]
    dir: PathBuf,
    #[arg(short = 'o', long = "output", value_name = "OUTPUT", required = true, value_parser = parse_xlsx_path)]
    output: PathBuf,
    #[arg(short = 'x', required = true)]
    x: usize,
    #[arg(short = 'y', required = true)]
    y: usize,
    #[arg(short = 'c', long = "cell_distance", required = true)]
    cd: f64,
}

#[derive(Serialize, Deserialize)]
struct F64Map {
    pub timestep: u32,
    pub width: u32,
    pub height: u32,
    pub data: Box<[f64]>,
}

#[derive(Serialize, Deserialize)]
struct InfectionStateMap {
    pub timestep: u32,
    pub width: u32,
    pub height: u32,
    pub healthy_sites: Box<[(u32, u32)]>,
    pub infected_sites: Box<[(u32, u32)]>,
    pub ignored_sites: Box<[(u32, u32)]>,
}

#[derive(Default)]
struct MapGrouping {
    foi: Option<F64Map>,
    infection: Option<InfectionStateMap>,
}

struct CombinedState {
    timestep: u32,
    width: u32,
    height: u32,
    foi_data: Box<[f64]>,
    healthy_sites: Box<[(u32, u32)]>,
    infected_sites: Box<[(u32, u32)]>,
    ignored_sites: Box<[(u32, u32)]>,
}

pub fn euclidean_distance(a: &(usize, usize), b: &(usize, usize)) -> f64 {
    (((b.0 - a.0).pow(2) + (b.1 - a.1).pow(2)) as f64).sqrt()
}

pub fn index_to_coordinates(index: usize, width: usize) -> (usize, usize) {
    (index % width, index / width)
}

pub fn coordinates_to_index(x: usize, y: usize, width: usize) -> usize {
    y * width + x
}

fn main() {
    let args = Args::parse();
    if let Err(err) = run(args) {
        eprintln!("{}", err);
        std::process::exit(1);
    }
}

fn run(args: Args) -> Result<(), Box<dyn Error>> {
    if !args.dir.is_dir() {
        return Err(format!("Not a directory: {}", args.dir.display()).into());
    }
    let foi_dir = args.dir.join("foi");
    let infection_dir = args.dir.join("infection");
    if !foi_dir.is_dir() {
        return Err(format!("Missing subdirectory: {}", foi_dir.display()).into());
    }
    if !infection_dir.is_dir() {
        return Err(format!("Missing subdirectory: {}", infection_dir.display()).into());
    }

    let mut foi_files: Vec<PathBuf> = WalkDir::new(&foi_dir)
        .into_iter()
        .filter_map(Result::ok)
        .filter(|e| e.file_type().is_file())
        .map(|e| e.path().to_path_buf())
        .filter(|p| p.extension().and_then(|s| s.to_str()) == Some("bin"))
        .filter(|p| {
            p.file_stem()
                .and_then(|s| s.to_str())
                .map(|s| s.chars().all(|c| c.is_ascii_digit()))
                .unwrap_or(false)
        })
        .collect();
    foi_files.sort_by(|a, b| compare_paths_natural(a, b));

    let mut infection_files: Vec<PathBuf> = WalkDir::new(&infection_dir)
        .into_iter()
        .filter_map(Result::ok)
        .filter(|e| e.file_type().is_file())
        .map(|e| e.path().to_path_buf())
        .filter(|p| p.extension().and_then(|s| s.to_str()) == Some("bin"))
        .filter(|p| {
            p.file_stem()
                .and_then(|s| s.to_str())
                .map(|s| s.chars().all(|c| c.is_ascii_digit()))
                .unwrap_or(false)
        })
        .collect();
    infection_files.sort_by(|a, b| compare_paths_natural(a, b));

    let mut by_timestep: HashMap<u32, MapGrouping> = HashMap::new();
    for path in foi_files {
        let bytes = fs::read(&path)?;
        match bincode::deserialize::<F64Map>(&bytes) {
            Ok(map) => {
                let ts = map.timestep;
                let entry = by_timestep.entry(ts).or_default();
                entry.foi = Some(map);
            }
            Err(err) => {
                eprintln!("ERR {} {}", path.display(), err);
            }
        }
    }
    for path in infection_files {
        let bytes = fs::read(&path)?;
        match bincode::deserialize::<InfectionStateMap>(&bytes) {
            Ok(map) => {
                let ts = map.timestep;
                let entry = by_timestep.entry(ts).or_default();
                entry.infection = Some(map);
            }
            Err(err) => {
                eprintln!("ERR {} {}", path.display(), err);
            }
        }
    }
    let mut combined: BTreeMap<u32, CombinedState> = BTreeMap::new();
    for (ts, group) in by_timestep.into_iter() {
        if let (Some(foi), Some(infection)) = (group.foi, group.infection) {
            if foi.timestep != ts || infection.timestep != ts {
                return Err(format!("Mismatched timestep for ts {}", ts).into());
            }
            if foi.width != infection.width || foi.height != infection.height {
                return Err(format!(
                    "Mismatched dimensions at ts {}: foi {}x{} vs infection {}x{}",
                    ts, foi.width, foi.height, infection.width, infection.height
                )
                .into());
            }
            combined.insert(
                ts,
                CombinedState {
                    timestep: ts,
                    width: foi.width,
                    height: foi.height,
                    foi_data: foi.data,
                    healthy_sites: infection.healthy_sites,
                    infected_sites: infection.infected_sites,
                    ignored_sites: infection.ignored_sites,
                },
            );
        }
    }
    {
        let (x, y, cd) = (args.x, args.y, args.cd);
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.write_string(0, 0, "timestep")?;
        worksheet.write_string(0, 1, "average_distance_from_source")?;
        worksheet.write_string(0, 2, "newly_infected_sites")?;
        worksheet.write_string(0, 3, "newly_infected_sites_change_ratio")?;
        worksheet.write_string(0, 4, "infected_area")?;
        worksheet.write_string(0, 5, "infection_area_change_ratio")?;
        worksheet.write_string(0, 6, "foi_99th_percentile")?;
        worksheet.write_string(0, 7, "foi_95th_percentile")?;
        let mut row: u32 = 1;
        let mut previous_number_of_infected_sites: usize = 0;
        let mut previous_infected_area: f64 = 0.0;
        let mut previous_newly_infected_sites: usize = 0;
        for (_, state) in combined.into_iter() {
            //percentile
            let (foi_99th_percentile, foi_95th_percentile) = {
                let mut foi_99th_percentile: TopPercentile<F64Ord, (usize, usize)> =
                    TopPercentile::new(0.99, state.foi_data.len());
                let mut foi_95th_percentile: TopPercentile<F64Ord, (usize, usize)> =
                    TopPercentile::new(0.95, state.foi_data.len());
                for (i, foi) in state.foi_data.iter().enumerate() {
                    foi_99th_percentile
                        .insert(F64Ord(*foi), index_to_coordinates(i, state.width as usize));
                    foi_95th_percentile
                        .insert(F64Ord(*foi), index_to_coordinates(i, state.width as usize));
                }
                let foi_99th_percentile = foi_99th_percentile
                    .smallest_primary()
                    .ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::Other,
                            "not enough values inserted to derive 99th percentile",
                        )
                    })?
                    .0
                    .0;
                let foi_95th_percentile = foi_95th_percentile
                    .smallest_primary()
                    .ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::Other,
                            "not enough values inserted to derive 95th percentile",
                        )
                    })?
                    .0
                    .0;
                (foi_99th_percentile, foi_95th_percentile)
            };

            //mean distance from source
            let infection_source: (usize, usize) = (x, y);
            let mut cumulative_distance_from_source: f64 = 0.0;
            let mut total_non_source_sites: usize = 0;
            for site in state.infected_sites.iter() {
                let site = (site.0 as usize, site.1 as usize);
                let distance = euclidean_distance(&infection_source, &site) * cd;
                if distance > 0.0 {
                    cumulative_distance_from_source += distance;
                    total_non_source_sites += 1;
                }
            }
            let mut mean_distance_from_source = if total_non_source_sites > 0 {
                cumulative_distance_from_source / total_non_source_sites as f64
            } else {
                0.0
            };
            if mean_distance_from_source.is_nan() {
                mean_distance_from_source = 0.0;
            }

            //newly infected sites
            let newly_infected_sites: usize =
                state.infected_sites.len() - previous_number_of_infected_sites;
            let newly_infected_sites_change_ratio = if previous_newly_infected_sites > 0 {
                newly_infected_sites as f64 / previous_newly_infected_sites as f64
            } else {
                1.0
            };
            previous_newly_infected_sites = newly_infected_sites;
            previous_number_of_infected_sites = state.infected_sites.len();

            //infected area
            let infected_area = state.infected_sites.len() as f64
                / (state.width as usize * state.height as usize) as f64;
            let mut infection_area_change_ratio = if previous_infected_area > 0.0 {
                infected_area / previous_infected_area
            } else {
                1.0
            };
            if infection_area_change_ratio.is_infinite() {
                infection_area_change_ratio = 1.0;
            }
            previous_infected_area = infected_area;

            //write to xlsx row
            worksheet.write_number(row, 0, state.timestep as f64)?;
            worksheet.write_number(row, 1, mean_distance_from_source)?;
            worksheet.write_number(row, 2, newly_infected_sites as f64)?;
            worksheet.write_number(row, 3, newly_infected_sites_change_ratio)?;
            worksheet.write_number(row, 4, infected_area)?;
            worksheet.write_number(row, 5, infection_area_change_ratio)?;
            worksheet.write_number(row, 6, foi_99th_percentile)?;
            worksheet.write_number(row, 7, foi_95th_percentile)?;
            row += 1;
        }
        workbook.save(&args.output)?;
    }

    Ok(())
}

fn compare_paths_natural(a: &PathBuf, b: &PathBuf) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    let sa = a.file_stem().and_then(|s| s.to_str()).unwrap_or("");
    let sb = b.file_stem().and_then(|s| s.to_str()).unwrap_or("");
    let na = trailing_number(sa);
    let nb = trailing_number(sb);
    match (na, nb) {
        (Some(x), Some(y)) => x.cmp(&y).then_with(|| sa.cmp(sb)),
        (Some(_), None) => Ordering::Less,
        (None, Some(_)) => Ordering::Greater,
        (None, None) => sa.cmp(sb),
    }
}

fn trailing_number(s: &str) -> Option<u64> {
    let mut i = s.len();
    while i > 0 && s.as_bytes()[i - 1].is_ascii_digit() {
        i -= 1;
    }
    if i == s.len() {
        return None;
    }
    s[i..].parse::<u64>().ok()
}
