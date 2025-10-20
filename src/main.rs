use clap::Parser;
use rust_xlsxwriter::{Color, ConditionalFormatCell, ConditionalFormatCellRule, Format, Workbook};
use serde::{Deserialize, Serialize};
use std::borrow::Borrow;
use std::collections::{BTreeMap, HashMap};
use std::error::Error;
use std::path::PathBuf;
use std::{fs, io};
use walkdir::WalkDir;

use std::cmp::Ordering;

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

fn mean_kahan<I>(iter: I) -> f64
where
    I: IntoIterator,
    I::Item: Borrow<f64>,
{
    let mut sum = 0.0f64;
    let mut c = 0.0f64; // compensation
    let mut n = 0usize;

    for v in iter {
        let x = *v.borrow(); // works for &f64 and f64
        let y = x - c;
        let t = sum + y;
        c = (t - sum) - y;
        sum = t;
        n += 1;
    }
    if n == 0 { 0.0 } else { sum / n as f64 }
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
                let timestep = map.timestep;
                let entry = by_timestep.entry(timestep).or_default();
                entry.infection = Some(map);
            }
            Err(err) => {
                eprintln!("ERR {} {}", path.display(), err);
            }
        }
    }
    let mut combined: BTreeMap<u32, CombinedState> = BTreeMap::new();
    for (timestep, group) in by_timestep.into_iter() {
        if let (Some(foi), Some(infection)) = (group.foi, group.infection) {
            if foi.timestep != timestep || infection.timestep != timestep {
                return Err(format!("Mismatched timestep for ts {}", timestep).into());
            }
            if foi.width != infection.width || foi.height != infection.height {
                return Err(format!(
                    "Mismatched dimensions at ts {}: foi {}x{} vs infection {}x{}",
                    timestep, foi.width, foi.height, infection.width, infection.height
                )
                .into());
            }
            combined.insert(
                timestep,
                CombinedState {
                    timestep,
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
        worksheet.write_string(0, 1, "inf_total")?; //total_number_of_infected_sites
        worksheet.write_string(0, 2, "inf_mean_dist_m")?; //infection_mean_distance_from_source
        worksheet.write_string(0, 3, "inf_new")?; //newly_infected_sites
        worksheet.write_string(0, 4, "inf_new_change")?; //newly_infected_sites_change_ratio
        worksheet.write_string(0, 5, "inf_new_change_mod")?; //newly_infected_sites_change_ratio
        worksheet.write_string(0, 6, "inf_area_m2")?; //infected_area
        worksheet.write_string(0, 7, "inf_area_change_m2")?;
        worksheet.write_string(0, 8, "inf_area_change_mod")?; //infection_area_change_ratio
        worksheet.write_string(0, 9, "foi_99th_p")?; //foi_99th_percentile
        worksheet.write_string(0, 10, "foi_95th_p")?; //foi_95th_percentile
        worksheet.write_string(0, 11, "foi_99th_p_mean_dist_m")?; //foi_99th_percentile_mean_distance_from_source
        worksheet.write_string(0, 12, "foi_95th_p_mean_dist_m")?; //foi_95th_percentile_mean_distance_from_source
        worksheet.write_string(0, 13, "inf_99th_p")?; //infection_99th_percentile
        worksheet.write_string(0, 14, "inf_95th_p")?; //infection_95th_percentile
        worksheet.write_string(0, 15, "inf_99th_p_mean_dist_m")?; //infection_mean_distance_of_99th_percentile_from_source
        worksheet.write_string(0, 16, "inf_95th_p_mean_dist_m")?; //infection_mean_distance_of_95th_percentile_from_source
        worksheet.write_string(0, 17, "inf_99th_p_spread_mpy")?; //infection spread rate in meters per year from 99th percentile
        worksheet.write_string(0, 18, "inf_95th_p_spread_mpy")?; //infection spread rate in meters per year from 95th percentile
        let mut row: u32 = 1;
        let mut previous_number_of_infected_sites: usize = 0;
        let mut previous_infected_area: f64 = 0.0;
        let mut previous_newly_infected_sites: usize = 0;
        let mut previous_infection_99th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_95th_percentile_mean_distance_from_source: f64 = 0.0;
        for (_, state) in combined.into_iter() {
            //validation
            if state.foi_data.iter().any(|v| v.is_nan()) {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "NaN encountered in foi_data",
                )
                .into());
            }

            let infection_source: (usize, usize) = (x, y);
            let infection_distances = {
                let mut map: HashMap<(usize, usize), f64> = HashMap::new();
                for &(x, y) in state.infected_sites.iter() {
                    let site = (x as usize, y as usize);
                    let distance = euclidean_distance(&infection_source, &site) * cd;
                    //exclude initial infection point
                    if distance > 0.0 {
                        map.insert(site, distance);
                    }
                }
                map
            };

            let foi_99th_percentile = percentile_nearest(&state.foi_data, 0.99)?;
            let foi_95th_percentile = percentile_nearest(&state.foi_data, 0.95)?;

            let infection_99th_percentile = {
                let values = infection_distances.values().copied().collect::<Vec<f64>>();
                if values.len() == 0 {
                    0.0
                } else {
                    percentile_nearest(&values, 0.99)?
                }
            };
            let infection_95th_percentile = {
                let values = infection_distances.values().copied().collect::<Vec<f64>>();
                if values.len() == 0 {
                    0.0
                } else {
                    percentile_nearest(&values, 0.95)?
                }
            };
            let infection_mean_distance_from_source = {
                let iter = infection_distances
                    .values()
                    .filter(|&distance| *distance > 0.0);
                mean_kahan(iter)
            };

            let infection_99th_percentile_mean_distance_from_source = {
                let iter = infection_distances
                    .values()
                    .filter(|&distance| *distance > 0.0 && *distance >= infection_99th_percentile);
                mean_kahan(iter)
            };

            let infection_95th_percentile_mean_distance_from_source = {
                let iter = infection_distances
                    .values()
                    .filter(|&distance| *distance > 0.0 && *distance >= infection_95th_percentile);
                mean_kahan(iter)
            };
            let (
                infection_99th_percentile_spread_rate_mpy,
                infection_95th_percentile_spread_rate_mpy,
            ) = {
                let mut infection_99th_percentile_spread_rate_mpy =
                    infection_99th_percentile_mean_distance_from_source
                        - previous_infection_99th_percentile_mean_distance_from_source;
                let mut infection_95th_percentile_spread_rate_mpy =
                    infection_95th_percentile_mean_distance_from_source
                        - previous_infection_95th_percentile_mean_distance_from_source;
                //if infection_99th_percentile_spread_rate_mpy.is_sign_negative() {
                println!(
                    "99: timestep: {timestep}, {infection_99th_percentile_mean_distance_from_source} - {previous_infection_99th_percentile_mean_distance_from_source} = {infection_99th_percentile_spread_rate_mpy}",
                    timestep = state.timestep,
                );
                //infection_99th_percentile_spread_rate_mpy = 0.0;
                //}
                //if infection_95th_percentile_spread_rate_mpy.is_sign_negative() {
                println!(
                    "95: timestep: {timestep}, {infection_95th_percentile_mean_distance_from_source} - {previous_infection_95th_percentile_mean_distance_from_source} = {infection_95th_percentile_spread_rate_mpy}",
                    timestep = state.timestep,
                );
                //infection_95th_percentile_spread_rate_mpy = 0.0;
                //}
                previous_infection_99th_percentile_mean_distance_from_source =
                    infection_99th_percentile_mean_distance_from_source;
                previous_infection_95th_percentile_mean_distance_from_source =
                    infection_95th_percentile_mean_distance_from_source;
                (
                    infection_99th_percentile_spread_rate_mpy,
                    infection_95th_percentile_spread_rate_mpy,
                )
            };

            let (foi_entries_above_99th_percentile, foi_entries_above_95th_percentile) = {
                let mut foi_entries_above_99th_percentile = Vec::new();
                let mut foi_entries_above_95th_percentile = Vec::new();
                for (index, foi) in state.foi_data.iter().enumerate() {
                    if foi >= &foi_99th_percentile || foi >= &foi_95th_percentile {
                        let coordinates = index_to_coordinates(index, state.width as usize);
                        let distance = euclidean_distance(&infection_source, &coordinates) * cd;
                        if distance == 0.0 {
                            continue;
                        }
                        if foi >= &foi_99th_percentile {
                            foi_entries_above_99th_percentile.push(distance);
                        }
                        if foi >= &foi_95th_percentile {
                            foi_entries_above_95th_percentile.push(distance);
                        }
                    }
                }
                (
                    foi_entries_above_99th_percentile,
                    foi_entries_above_95th_percentile,
                )
            };

            let foi_99th_percentile_mean_distance_from_source =
                mean_kahan(foi_entries_above_99th_percentile.iter());
            let foi_95th_percentile_mean_distance_from_source =
                mean_kahan(foi_entries_above_95th_percentile.iter());

            //newly infected sites
            let newly_infected_sites: usize =
                state.infected_sites.len() - previous_number_of_infected_sites;
            let (newly_infected_sites_change_modifier, newly_infected_sites_change) =
                if previous_newly_infected_sites > 0 {
                    (
                        newly_infected_sites as f64 / previous_newly_infected_sites as f64,
                        newly_infected_sites as f64 - previous_newly_infected_sites as f64,
                    )
                } else {
                    (1.0, 0.0)
                };
            previous_newly_infected_sites = newly_infected_sites;
            previous_number_of_infected_sites = state.infected_sites.len();

            let total_number_of_infected_sites = state.infected_sites.len();

            //infected area
            let infected_area = state.infected_sites.len() as f64
                / (state.width as usize * state.height as usize) as f64;
            let (infection_area_change_modifier, infection_area_change) =
                if previous_infected_area > 0.0 {
                    (
                        infected_area / previous_infected_area,
                        infected_area - previous_infected_area,
                    )
                } else {
                    (1.0, 0.0)
                };
            if infection_area_change_modifier.is_infinite() {
                panic!("infection_area_change_modifier is infinite");
            }
            previous_infected_area = infected_area;

            //write to xlsx row
            worksheet.write_number(row, 0, state.timestep as f64)?;
            worksheet.write_number(row, 1, total_number_of_infected_sites as f64)?;
            worksheet.write_number(row, 2, infection_mean_distance_from_source)?;
            worksheet.write_number(row, 3, newly_infected_sites as f64)?;
            worksheet.write_number(row, 4, newly_infected_sites_change)?;
            worksheet.write_number(row, 5, newly_infected_sites_change_modifier)?;
            worksheet.write_number(row, 6, infected_area)?;
            worksheet.write_number(row, 7, infection_area_change)?;
            worksheet.write_number(row, 8, infection_area_change_modifier)?;
            worksheet.write_number(row, 9, foi_99th_percentile)?;
            worksheet.write_number(row, 10, foi_95th_percentile)?;
            worksheet.write_number(row, 11, foi_99th_percentile_mean_distance_from_source)?;
            worksheet.write_number(row, 12, foi_95th_percentile_mean_distance_from_source)?;
            worksheet.write_number(row, 13, infection_99th_percentile)?;
            worksheet.write_number(row, 14, infection_95th_percentile)?;
            worksheet.write_number(row, 15, infection_99th_percentile_mean_distance_from_source)?;
            worksheet.write_number(row, 16, infection_95th_percentile_mean_distance_from_source)?;
            worksheet.write_number(row, 17, infection_99th_percentile_spread_rate_mpy)?;
            worksheet.write_number(row, 18, infection_95th_percentile_spread_rate_mpy)?;
            row += 1;
        }

        let red_format = Format::new().set_background_color(Color::RGB(0xFFC7CE));
        let green_format = Format::new().set_background_color(Color::RGB(0xC6EFCE));
        let neutral_format = Format::new().set_background_color(Color::RGB(0xFFEB9C));

        let negative_condition = ConditionalFormatCell::new()
            .set_rule(ConditionalFormatCellRule::LessThan(0.0))
            .set_format(red_format);

        let positive_condition = ConditionalFormatCell::new()
            .set_rule(ConditionalFormatCellRule::GreaterThan(0.0))
            .set_format(green_format);

        let zero_condition = ConditionalFormatCell::new()
            .set_rule(ConditionalFormatCellRule::EqualTo(0.0))
            .set_format(neutral_format);

        //inf_new_change
        worksheet.add_conditional_format(1, 4, row - 1, 4, &negative_condition)?;
        worksheet.add_conditional_format(1, 4, row - 1, 4, &positive_condition)?;
        worksheet.add_conditional_format(1, 4, row - 1, 4, &zero_condition)?;

        //inf_area_change
        //worksheet.add_conditional_format(1, 7, row - 1, 7, &negative_condition)?;
        //worksheet.add_conditional_format(1, 7, row - 1, 7, &positive_condition)?;
        //worksheet.add_conditional_format(1, 7, row - 1, 7, &zero_condition)?;

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

fn percentile_nearest(data: &[f64], q: f64) -> Result<f64, Box<dyn Error>> {
    //validation
    if data.is_empty() {
        return Err(io::Error::new(io::ErrorKind::InvalidData, "empty data").into());
    }
    if !(0.0..=1.0).contains(&q) {
        return Err(io::Error::new(io::ErrorKind::InvalidInput, "q out of range").into());
    }

    //calculate percentile
    let mut buf: Vec<f64> = data.to_vec();
    let n = buf.len();
    let idx = if q >= 1.0 {
        n - 1
    } else {
        (q * (n as f64 - 1.0)).round() as usize
    };
    let (_, nth, _) =
        buf.select_nth_unstable_by(idx, |a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    Ok(*nth)
}
