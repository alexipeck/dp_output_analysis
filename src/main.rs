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

mod image_grid;
use crate::image_grid::{GridRenderConfig, render_foi_png_gray16, render_infection_state_png};

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
    #[arg(
        short = 'o',
        long = "output-dir",
        value_name = "OUTPUT_DIR",
        required = true
    )]
    output_dir: PathBuf,
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
    pub healthy_biomass: Box<[u64]>,
    pub infected_biomass: Box<[u64]>,
    pub ignored_biomass: Box<[u64]>,
}

#[derive(Serialize, Deserialize)]
struct MortalityMap {
    pub timestep: u32,
    pub width: u32,
    pub height: u32,
    pub data: Box<[u32]>,
}

#[derive(Default)]
struct MapGrouping {
    foi: Option<F64Map>,
    infection: Option<InfectionStateMap>,
    mortality: Option<MortalityMap>,
}

struct Foi {
    data: Box<[f64]>,
}

struct Infection {
    healthy_sites: Box<[(u32, u32)]>,
    infected_sites: Box<[(u32, u32)]>,
    ignored_sites: Box<[(u32, u32)]>,
    healthy_biomass: Box<[u64]>,
    infected_biomass: Box<[u64]>,
    ignored_biomass: Box<[u64]>,
}

struct Mortality {
    data: Box<[u32]>,
}

struct CombinedState {
    timestep: u32,
    width: u32,
    height: u32,
    foi: Option<Foi>,
    infection: Option<Infection>,
    mortality: Option<Mortality>,
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
    let mut compensation = 0.0f64;
    let mut n = 0usize;

    for value in iter {
        let x = *value.borrow(); // works for &f64 and f64
        let y = x - compensation;
        let t = sum + y;
        compensation = (t - sum) - y;
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
    if !args.output_dir.exists() {
        fs::create_dir_all(&args.output_dir)?;
    }
    let foi_dir = args.dir.join("foi");
    let infection_dir = args.dir.join("infection");
    let mortality_dir = args.dir.join("mortality");
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

    let mut mortality_files: Vec<PathBuf> = WalkDir::new(&mortality_dir)
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
    mortality_files.sort_by(|a, b| compare_paths_natural(a, b));

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
                let entry = by_timestep.entry(map.timestep).or_default();
                entry.infection = Some(map);
            }
            Err(err) => {
                eprintln!("ERR {} {}", path.display(), err);
            }
        }
    }
    for path in mortality_files {
        let bytes = fs::read(&path)?;
        match bincode::deserialize::<MortalityMap>(&bytes) {
            Ok(map) => {
                let timestep = map.timestep;
                let entry = by_timestep.entry(timestep).or_default();
                entry.mortality = Some(map);
            }
            Err(err) => {
                eprintln!("ERR {} {}", path.display(), err);
            }
        }
    }
    let mut combined: BTreeMap<u32, CombinedState> = BTreeMap::new();
    for (timestep, group) in by_timestep.into_iter() {
        let mut combined_state = CombinedState {
            timestep,
            width: 0,
            height: 0,
            foi: None,
            infection: None,
            mortality: None,
        };

        if let Some(foi) = group.foi {
            if foi.timestep != timestep {
                return Err(format!("Mismatched timestep for foi at ts {}", timestep).into());
            }
            combined_state.width = foi.width;
            combined_state.height = foi.height;
            combined_state.foi = Some(Foi { data: foi.data });
        }

        if let Some(infection) = group.infection {
            if infection.timestep != timestep {
                return Err(format!("Mismatched timestep for infection at ts {}", timestep).into());
            }
            if combined_state.width != 0
                && (combined_state.width != infection.width
                    || combined_state.height != infection.height)
            {
                return Err(format!(
                    "Mismatched dimensions at ts {}: existing {}x{} vs infection {}x{}",
                    timestep,
                    combined_state.width,
                    combined_state.height,
                    infection.width,
                    infection.height
                )
                .into());
            }
            if combined_state.width == 0 {
                combined_state.width = infection.width;
                combined_state.height = infection.height;
            }
            if infection.healthy_biomass.len()
                != combined_state.width as usize * combined_state.height as usize
            {
                return Err(format!(
                    "Mismatched healthy_biomass length at ts {}: {}",
                    timestep,
                    infection.healthy_biomass.len()
                )
                .into());
            }
            if infection.infected_biomass.len() != infection.healthy_biomass.len() {
                return Err(format!(
                    "Mismatched infected_biomass length at ts {}: {}",
                    timestep,
                    infection.infected_biomass.len()
                )
                .into());
            }
            if infection.ignored_biomass.len() != infection.healthy_biomass.len() {
                return Err(format!(
                    "Mismatched ignored_biomass length at ts {}: {}",
                    timestep,
                    infection.ignored_biomass.len()
                )
                .into());
            }
            combined_state.infection = Some(Infection {
                healthy_sites: infection.healthy_sites,
                infected_sites: infection.infected_sites,
                ignored_sites: infection.ignored_sites,
                healthy_biomass: infection.healthy_biomass,
                infected_biomass: infection.infected_biomass,
                ignored_biomass: infection.ignored_biomass,
            });
        }

        if let Some(mortality) = group.mortality {
            if mortality.timestep != timestep {
                return Err(format!("Mismatched timestep for mortality at ts {}", timestep).into());
            }
            if combined_state.width != 0
                && (combined_state.width != mortality.width
                    || combined_state.height != mortality.height)
            {
                return Err(format!(
                    "Mismatched dimensions at ts {}: existing {}x{} vs mortality {}x{}",
                    timestep,
                    combined_state.width,
                    combined_state.height,
                    mortality.width,
                    mortality.height
                )
                .into());
            }
            if combined_state.width == 0 {
                combined_state.width = mortality.width;
                combined_state.height = mortality.height;
            }
            combined_state.mortality = Some(Mortality {
                data: mortality.data,
            });
        }

        if combined_state.width == 0 {
            return Err(format!("No data available for timestep {}", timestep).into());
        }

        combined.insert(timestep, combined_state);
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
        worksheet.write_string(0, 19, "t_annual_mortality")?; //annual mortality
        worksheet.write_string(0, 20, "t_hea_biomass")?; //total healthy biomass
        worksheet.write_string(0, 21, "t_inf_biomass")?; //total infected biomass
        worksheet.write_string(0, 22, "t_ign_biomass")?; //total ignored biomass
        worksheet.write_string(0, 23, "t_biomass")?; //total biomass
        worksheet.write_string(0, 24, "prop_inf_biomass")?; //proportion of infected biomass
        worksheet.write_string(0, 25, "prop_host_inf_biomass")?; //proportion of host infected biomass
        worksheet.write_string(0, 26, "t_hea_biomass_change")?; //total healthy biomass change
        worksheet.write_string(0, 27, "t_inf_biomass_change")?; //total infected biomass change
        worksheet.write_string(0, 28, "t_ign_biomass_change")?; //total ignored biomass change
        worksheet.write_string(0, 29, "t_biomass_change")?; //total biomass change
        worksheet.write_string(0, 30, "t_hea_biomass_change_mod")?; //total healthy biomass change modifier
        worksheet.write_string(0, 31, "t_inf_biomass_change_mod")?; //total infected biomass change modifier
        worksheet.write_string(0, 32, "t_ign_biomass_change_mod")?; //total ignored biomass change modifier
        worksheet.write_string(0, 33, "t_biomass_change_mod")?; //total biomass change modifier

        let mut row: u32 = 1;
        let mut previous_number_of_infected_sites: usize = 0;
        let mut previous_infected_area: f64 = 0.0;
        let mut previous_newly_infected_sites: usize = 0;
        let mut previous_infection_99th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_95th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_total_healthy_biomass: f64 = 0.0;
        let mut previous_total_infected_biomass: f64 = 0.0;
        let mut previous_total_ignored_biomass: f64 = 0.0;
        let mut previous_total_biomass: f64 = 0.0;
        let mut mortality_values: Vec<f64> = Vec::new();
        let mut img_time_sum_ms: u128 = 0;
        let mut img_time_count: usize = 0;

        let mut global_foi_min = f64::INFINITY;
        let mut global_foi_max = f64::NEG_INFINITY;
        for (_, state) in combined.iter() {
            if let Some(foi) = &state.foi {
                for &v in foi.data.iter() {
                    if v.is_finite() {
                        if v < global_foi_min {
                            global_foi_min = v;
                        }
                        if v > global_foi_max {
                            global_foi_max = v;
                        }
                    }
                }
            }
        }

        for (_, state) in combined.into_iter() {
            let infection_source: (usize, usize) = (x, y);

            if let Some(foi) = &state.foi {
                if foi.data.iter().any(|v| v.is_nan()) {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "NaN encountered in foi_data",
                    )
                    .into());
                }

                let foi_99th_percentile = percentile_nearest(&foi.data, 0.99)?;
                let foi_95th_percentile = percentile_nearest(&foi.data, 0.95)?;

                let (foi_entries_above_99th_percentile, foi_entries_above_95th_percentile) = {
                    let mut foi_entries_above_99th_percentile = Vec::new();
                    let mut foi_entries_above_95th_percentile = Vec::new();
                    for (index, foi_value) in foi.data.iter().enumerate() {
                        if foi_value >= &foi_99th_percentile || foi_value >= &foi_95th_percentile {
                            let coordinates = index_to_coordinates(index, state.width as usize);
                            let distance = euclidean_distance(&infection_source, &coordinates) * cd;
                            if distance == 0.0 {
                                continue;
                            }
                            if foi_value >= &foi_99th_percentile {
                                foi_entries_above_99th_percentile.push(distance);
                            }
                            if foi_value >= &foi_95th_percentile {
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

                worksheet.write_number(row, 9, foi_99th_percentile)?;
                worksheet.write_number(row, 10, foi_95th_percentile)?;
                worksheet.write_number(row, 11, foi_99th_percentile_mean_distance_from_source)?;
                worksheet.write_number(row, 12, foi_95th_percentile_mean_distance_from_source)?;

                let cfg = GridRenderConfig::default();
                let normalized: Vec<f64> = foi
                    .data
                    .iter()
                    .map(|&v| {
                        if v.is_finite() && global_foi_max > global_foi_min {
                            let mut n = (v - global_foi_min) / (global_foi_max - global_foi_min);
                            if n < 0.0 {
                                n = 0.0;
                            }
                            if n > 1.0 {
                                n = 1.0;
                            }
                            n
                        } else {
                            0.0
                        }
                    })
                    .collect();
                let _ = render_foi_png_gray16(
                    &args.output_dir,
                    state.timestep,
                    state.width,
                    state.height,
                    &normalized,
                    &cfg,
                )?;
            }

            if let Some(infection) = &state.infection {
                let infection_distances = {
                    let mut map: HashMap<(usize, usize), f64> = HashMap::new();
                    for &(x, y) in infection.infected_sites.iter() {
                        let site = (x as usize, y as usize);
                        let distance = euclidean_distance(&infection_source, &site) * cd;
                        //exclude initial infection point
                        if distance > 0.0 {
                            map.insert(site, distance);
                        }
                    }
                    map
                };

                let total_number_of_infected_sites = infection.infected_sites.len();
                let newly_infected_sites =
                    infection.infected_sites.len() - previous_number_of_infected_sites;
                let infected_area = infection.infected_sites.len() as f64
                    / (state.width as usize * state.height as usize) as f64;

                let values = infection_distances.values().copied().collect::<Vec<f64>>();
                let infection_99th_percentile = if values.len() > 0 {
                    percentile_nearest(&values, 0.99)?
                } else {
                    0.0
                };
                let infection_95th_percentile = if values.len() > 0 {
                    percentile_nearest(&values, 0.95)?
                } else {
                    0.0
                };

                let infection_mean_distance_from_source = {
                    let iter = infection_distances
                        .values()
                        .filter(|&distance| *distance > 0.0);
                    mean_kahan(iter)
                };

                let infection_99th_percentile_mean_distance_from_source = {
                    let iter = infection_distances.values().filter(|&distance| {
                        *distance > 0.0 && *distance >= infection_99th_percentile
                    });
                    mean_kahan(iter)
                };

                let infection_95th_percentile_mean_distance_from_source = {
                    let iter = infection_distances.values().filter(|&distance| {
                        *distance > 0.0 && *distance >= infection_95th_percentile
                    });
                    mean_kahan(iter)
                };

                let infection_99th_percentile_spread_rate_mpy =
                    infection_99th_percentile_mean_distance_from_source
                        - previous_infection_99th_percentile_mean_distance_from_source;
                let infection_95th_percentile_spread_rate_mpy =
                    infection_95th_percentile_mean_distance_from_source
                        - previous_infection_95th_percentile_mean_distance_from_source;
                previous_infection_99th_percentile_mean_distance_from_source =
                    infection_99th_percentile_mean_distance_from_source;
                previous_infection_95th_percentile_mean_distance_from_source =
                    infection_95th_percentile_mean_distance_from_source;

                let total_healthy_biomass = infection.healthy_biomass.iter().sum::<u64>() as f64;
                let total_infected_biomass = infection.infected_biomass.iter().sum::<u64>() as f64;
                let total_ignored_biomass = infection.ignored_biomass.iter().sum::<u64>() as f64;
                let total_biomass =
                    total_healthy_biomass + total_infected_biomass + total_ignored_biomass;

                let proportion_infected_biomass = total_infected_biomass / total_biomass;
                let proportion_host_infected_biomass =
                    total_infected_biomass / (total_healthy_biomass + total_infected_biomass);

                let (newly_infected_sites_change_modifier, newly_infected_sites_change) =
                    if previous_newly_infected_sites > 0 {
                        (
                            newly_infected_sites as f64 / previous_newly_infected_sites as f64,
                            newly_infected_sites as f64 - previous_newly_infected_sites as f64,
                        )
                    } else {
                        (1.0, 0.0)
                    };

                let (infection_area_change_modifier, infection_area_change) =
                    if previous_infected_area > 0.0 {
                        (
                            infected_area / previous_infected_area,
                            infected_area - previous_infected_area,
                        )
                    } else {
                        (1.0, 0.0)
                    };

                let (total_healthy_biomass_change_modifier, total_healthy_biomass_change) =
                    if previous_total_healthy_biomass > 0.0 {
                        (
                            total_healthy_biomass / previous_total_healthy_biomass,
                            total_healthy_biomass - previous_total_healthy_biomass,
                        )
                    } else {
                        (1.0, 0.0)
                    };
                let (total_infected_biomass_change_modifier, total_infected_biomass_change) =
                    if previous_total_infected_biomass > 0.0 {
                        (
                            total_infected_biomass / previous_total_infected_biomass,
                            total_infected_biomass - previous_total_infected_biomass,
                        )
                    } else {
                        (1.0, 0.0)
                    };
                let (total_ignored_biomass_change_modifier, total_ignored_biomass_change) =
                    if previous_total_ignored_biomass > 0.0 {
                        (
                            total_ignored_biomass / previous_total_ignored_biomass,
                            total_ignored_biomass - previous_total_ignored_biomass,
                        )
                    } else {
                        (1.0, 0.0)
                    };
                let (total_biomass_change_modifier, total_biomass_change) =
                    if previous_total_biomass > 0.0 {
                        (
                            total_biomass / previous_total_biomass,
                            total_biomass - previous_total_biomass,
                        )
                    } else {
                        (1.0, 0.0)
                    };

                previous_number_of_infected_sites = infection.infected_sites.len();
                previous_infected_area = infected_area;
                previous_newly_infected_sites = newly_infected_sites;
                previous_total_healthy_biomass = total_healthy_biomass;
                previous_total_infected_biomass = total_infected_biomass;
                previous_total_ignored_biomass = total_ignored_biomass;
                previous_total_biomass = total_biomass;

                worksheet.write_number(row, 1, total_number_of_infected_sites as f64)?;
                worksheet.write_number(row, 2, infection_mean_distance_from_source)?;
                worksheet.write_number(row, 3, newly_infected_sites as f64)?;
                worksheet.write_number(row, 4, newly_infected_sites_change)?;
                worksheet.write_number(row, 5, newly_infected_sites_change_modifier)?;
                worksheet.write_number(row, 6, infected_area)?;
                worksheet.write_number(row, 7, infection_area_change)?;
                worksheet.write_number(row, 8, infection_area_change_modifier)?;
                worksheet.write_number(row, 13, infection_99th_percentile)?;
                worksheet.write_number(row, 14, infection_95th_percentile)?;
                worksheet.write_number(
                    row,
                    15,
                    infection_99th_percentile_mean_distance_from_source,
                )?;
                worksheet.write_number(
                    row,
                    16,
                    infection_95th_percentile_mean_distance_from_source,
                )?;
                worksheet.write_number(row, 17, infection_99th_percentile_spread_rate_mpy)?;
                worksheet.write_number(row, 18, infection_95th_percentile_spread_rate_mpy)?;
                worksheet.write_number(row, 20, total_healthy_biomass)?;
                worksheet.write_number(row, 21, total_infected_biomass)?;
                worksheet.write_number(row, 22, total_ignored_biomass)?;
                worksheet.write_number(row, 23, total_biomass)?;
                worksheet.write_number(row, 24, proportion_infected_biomass)?;
                worksheet.write_number(row, 25, proportion_host_infected_biomass)?;
                worksheet.write_number(row, 26, total_healthy_biomass_change)?;
                worksheet.write_number(row, 27, total_infected_biomass_change)?;
                worksheet.write_number(row, 28, total_ignored_biomass_change)?;
                worksheet.write_number(row, 29, total_biomass_change)?;
                worksheet.write_number(row, 30, total_healthy_biomass_change_modifier)?;
                worksheet.write_number(row, 31, total_infected_biomass_change_modifier)?;
                worksheet.write_number(row, 32, total_ignored_biomass_change_modifier)?;
                worksheet.write_number(row, 33, total_biomass_change_modifier)?;
            }

            if let Some(mortality) = &state.mortality {
                let total_annual_mortality = mortality.data.iter().sum::<u32>() as f64;
                mortality_values.push(total_annual_mortality);

                worksheet.write_number(row, 19, total_annual_mortality)?;
            }

            if let Some(infection) = &state.infection {
                let cfg = GridRenderConfig::default();
                let start = std::time::Instant::now();
                let _ = render_infection_state_png(
                    &args.output_dir,
                    state.timestep,
                    state.width,
                    state.height,
                    &infection.healthy_sites,
                    &infection.infected_sites,
                    &infection.ignored_sites,
                    &cfg,
                )?;
                let ms = start.elapsed().as_millis();
                img_time_sum_ms += ms;
                img_time_count += 1;
            }

            worksheet.write_number(row, 0, state.timestep as f64)?;
            row += 1;
        }

        if img_time_count > 0 {
            let avg = img_time_sum_ms as f64 / img_time_count as f64;
            println!(
                "avg infection image render: {:.2} ms over {} timesteps",
                avg, img_time_count
            );
        }
        let red_format = Format::new().set_background_color(Color::RGB(0xFFC7CE));
        let green_format = Format::new().set_background_color(Color::RGB(0xC6EFCE));
        let neutral_format = Format::new().set_background_color(Color::RGB(0xFFEB9C));

        let negative_condition = ConditionalFormatCell::new()
            .set_rule(ConditionalFormatCellRule::LessThan(0.0))
            .set_format(&red_format);

        let positive_condition = ConditionalFormatCell::new()
            .set_rule(ConditionalFormatCellRule::GreaterThan(0.0))
            .set_format(&green_format);

        let zero_condition = ConditionalFormatCell::new()
            .set_rule(ConditionalFormatCellRule::EqualTo(0.0))
            .set_format(&neutral_format);

        let modifier_less_than_one_condition = ConditionalFormatCell::new()
            .set_rule(ConditionalFormatCellRule::LessThan(1.0))
            .set_format(&red_format);

        let modifier_greater_than_one_condition = ConditionalFormatCell::new()
            .set_rule(ConditionalFormatCellRule::GreaterThan(1.0))
            .set_format(&green_format);

        let modifier_equal_one_condition = ConditionalFormatCell::new()
            .set_rule(ConditionalFormatCellRule::EqualTo(1.0))
            .set_format(&neutral_format);

        //inf_new_change
        worksheet.add_conditional_format(1, 4, row - 1, 4, &negative_condition)?;
        worksheet.add_conditional_format(1, 4, row - 1, 4, &positive_condition)?;
        worksheet.add_conditional_format(1, 4, row - 1, 4, &zero_condition)?;

        //total_biomass_change
        worksheet.add_conditional_format(1, 29, row - 1, 29, &negative_condition)?;
        worksheet.add_conditional_format(1, 29, row - 1, 29, &positive_condition)?;
        worksheet.add_conditional_format(1, 29, row - 1, 29, &zero_condition)?;

        //total_infected_biomass_change
        worksheet.add_conditional_format(1, 27, row - 1, 27, &negative_condition)?;
        worksheet.add_conditional_format(1, 27, row - 1, 27, &positive_condition)?;
        worksheet.add_conditional_format(1, 27, row - 1, 27, &zero_condition)?;

        //total_healthy_biomass_change
        worksheet.add_conditional_format(1, 26, row - 1, 26, &negative_condition)?;
        worksheet.add_conditional_format(1, 26, row - 1, 26, &positive_condition)?;
        worksheet.add_conditional_format(1, 26, row - 1, 26, &zero_condition)?;

        //total_ignored_biomass_change
        worksheet.add_conditional_format(1, 28, row - 1, 28, &negative_condition)?;
        worksheet.add_conditional_format(1, 28, row - 1, 28, &positive_condition)?;
        worksheet.add_conditional_format(1, 28, row - 1, 28, &zero_condition)?;

        //total_healthy_biomass_change_modifier
        worksheet.add_conditional_format(1, 30, row - 1, 30, &modifier_less_than_one_condition)?;
        worksheet.add_conditional_format(
            1,
            30,
            row - 1,
            30,
            &modifier_greater_than_one_condition,
        )?;
        worksheet.add_conditional_format(1, 30, row - 1, 30, &modifier_equal_one_condition)?;

        //total_infected_biomass_change_modifier
        worksheet.add_conditional_format(1, 31, row - 1, 31, &modifier_less_than_one_condition)?;
        worksheet.add_conditional_format(
            1,
            31,
            row - 1,
            31,
            &modifier_greater_than_one_condition,
        )?;
        worksheet.add_conditional_format(1, 31, row - 1, 31, &modifier_equal_one_condition)?;

        //total_ignored_biomass_change_modifier
        worksheet.add_conditional_format(1, 32, row - 1, 32, &modifier_less_than_one_condition)?;
        worksheet.add_conditional_format(
            1,
            32,
            row - 1,
            32,
            &modifier_greater_than_one_condition,
        )?;
        worksheet.add_conditional_format(1, 32, row - 1, 32, &modifier_equal_one_condition)?;

        //total_biomass_change_modifier
        worksheet.add_conditional_format(1, 33, row - 1, 33, &modifier_less_than_one_condition)?;
        worksheet.add_conditional_format(
            1,
            33,
            row - 1,
            33,
            &modifier_greater_than_one_condition,
        )?;
        worksheet.add_conditional_format(1, 33, row - 1, 33, &modifier_equal_one_condition)?;

        //inf_area_change
        //worksheet.add_conditional_format(1, 7, row - 1, 7, &negative_condition)?;
        //worksheet.add_conditional_format(1, 7, row - 1, 7, &positive_condition)?;
        //worksheet.add_conditional_format(1, 7, row - 1, 7, &zero_condition)?;

        // Conditional formatting for total_annual_mortality column
        if !mortality_values.is_empty() {
            let min_mortality = mortality_values
                .iter()
                .fold(f64::INFINITY, |a, &b| a.min(b));
            let max_mortality = mortality_values
                .iter()
                .fold(f64::NEG_INFINITY, |a, &b| a.max(b));
            let median_mortality = percentile_nearest(&mortality_values, 0.5).unwrap_or(0.0);

            let mortality_min_format = Format::new().set_background_color(Color::RGB(0xC6EFCE));
            let mortality_max_format = Format::new().set_background_color(Color::RGB(0xFFC7CE));
            let mortality_median_format = Format::new().set_background_color(Color::RGB(0xFFEB9C));

            let mortality_min_condition = ConditionalFormatCell::new()
                .set_rule(ConditionalFormatCellRule::EqualTo(min_mortality))
                .set_format(&mortality_min_format);

            let mortality_max_condition = ConditionalFormatCell::new()
                .set_rule(ConditionalFormatCellRule::EqualTo(max_mortality))
                .set_format(&mortality_max_format);

            let mortality_median_condition = ConditionalFormatCell::new()
                .set_rule(ConditionalFormatCellRule::EqualTo(median_mortality))
                .set_format(&mortality_median_format);

            worksheet.add_conditional_format(1, 19, row - 1, 19, &mortality_min_condition)?;
            worksheet.add_conditional_format(1, 19, row - 1, 19, &mortality_max_condition)?;
            worksheet.add_conditional_format(1, 19, row - 1, 19, &mortality_median_condition)?;
        }

        let xlsx_path = args.output_dir.join("output.xlsx");
        workbook.save(&xlsx_path)?;
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
