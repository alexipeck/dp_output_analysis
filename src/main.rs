use clap::Parser;
use rust_xlsxwriter::worksheet::Worksheet;
use rust_xlsxwriter::{
    Color, ConditionalFormatCell, ConditionalFormatCellRule, Format, Note, XlsxError,
};
use serde::{Deserialize, Serialize};
use std::borrow::Borrow;
use std::collections::{BTreeMap, HashMap};
use std::error::Error;
use std::path::PathBuf;
use std::{fs, io};
use strum::{EnumIter, IntoEnumIterator};
use walkdir::WalkDir;

use lazy_static::lazy_static;
use std::cmp::Ordering;
use std::sync::atomic::{AtomicU32, Ordering as AtomicOrdering};

lazy_static! {
    static ref ROW_COUNTER: AtomicU32 = AtomicU32::new(1);
}

struct Workbook {
    workbook: rust_xlsxwriter::Workbook,
}

impl Workbook {
    fn new() -> Self {
        let mut workbook = rust_xlsxwriter::Workbook::new();
        workbook.add_worksheet();
        Self { workbook }
    }

    fn get_worksheet(&mut self) -> &mut Worksheet {
        self.workbook
            .worksheet_from_index(0)
            .expect("Worksheet not found")
    }

    fn add_conditional_format(
        &mut self,
        column: u16,
        condition: &ConditionalFormatCell,
    ) -> Result<(), XlsxError> {
        let row = ROW_COUNTER.load(AtomicOrdering::SeqCst);
        let worksheet = self.get_worksheet();
        worksheet.add_conditional_format(1, column, row - 1, column, condition)?;
        Ok(())
    }

    fn write_number(&mut self, column: u16, value: f64) -> Result<(), XlsxError> {
        let row = ROW_COUNTER.load(AtomicOrdering::SeqCst);
        let worksheet = self.get_worksheet();
        worksheet.write_number(row, column, value)?;
        Ok(())
    }

    fn save(&mut self, path: &str) -> Result<(), XlsxError> {
        self.workbook.save(path)
    }
}

enum ConditionalFormatter {
    ChangeModifier,
    ChangePercentage,
    ChangeAbsolute,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumIter)]
enum Column {
    Timestep,
    InfTotal,
    InfMeanDistMeters,
    InfNew,
    InfectedNewChange,
    InfNewChangePercentage,
    InfNewChangeMod,
    InfAreaSquareMeters,
    InfAreaChangeSquareMeters,
    InfAreaChangeSquareMetersPercentage,
    InfectedAreaChangeSquareMetersMod,
    Inf99thPercentile,
    Inf95thPercentile,
    Inf90thPercentile,
    Inf85thPercentile,
    Inf80thPercentile,
    Inf75thPercentile,
    Inf70thPercentile,
    Inf65thPercentile,
    Inf60thPercentile,
    Inf55thPercentile,
    Inf50thPercentile,
    Inf45thPercentile,
    Inf40thPercentile,
    Inf35thPercentile,
    Inf30thPercentile,
    Inf25thPercentile,
    Inf20thPercentile,
    Inf15thPercentile,
    Inf10thPercentile,
    Inf99thPercentileMeanDistanceMeters,
    Inf95thPercentileMeanDistanceMeters,
    Inf90thPercentileMeanDistanceMeters,
    Inf85thPercentileMeanDistanceMeters,
    Inf80thPercentileMeanDistanceMeters,
    Inf75thPercentileMeanDistanceMeters,
    Inf70thPercentileMeanDistanceMeters,
    Inf65thPercentileMeanDistanceMeters,
    Inf60thPercentileMeanDistanceMeters,
    Inf55thPercentileMeanDistanceMeters,
    Inf50thPercentileMeanDistanceMeters,
    Inf45thPercentileMeanDistanceMeters,
    Inf40thPercentileMeanDistanceMeters,
    Inf35thPercentileMeanDistanceMeters,
    Inf30thPercentileMeanDistanceMeters,
    Inf25thPercentileMeanDistanceMeters,
    Inf20thPercentileMeanDistanceMeters,
    Inf15thPercentileMeanDistanceMeters,
    Inf10thPercentileMeanDistanceMeters,
    Inf99thPercentileSpreadMetersPerYear,
    Inf95thPercentileSpreadMetersPerYear,
    Inf90thPercentileSpreadMetersPerYear,
    Inf85thPercentileSpreadMetersPerYear,
    Inf80thPercentileSpreadMetersPerYear,
    Inf75thPercentileSpreadMetersPerYear,
    Inf70thPercentileSpreadMetersPerYear,
    Inf65thPercentileSpreadMetersPerYear,
    Inf60thPercentileSpreadMetersPerYear,
    Inf55thPercentileSpreadMetersPerYear,
    Inf50thPercentileSpreadMetersPerYear,
    Inf45thPercentileSpreadMetersPerYear,
    Inf40thPercentileSpreadMetersPerYear,
    Inf35thPercentileSpreadMetersPerYear,
    Inf30thPercentileSpreadMetersPerYear,
    Inf25thPercentileSpreadMetersPerYear,
    Inf20thPercentileSpreadMetersPerYear,
    Inf15thPercentileSpreadMetersPerYear,
    Inf10thPercentileSpreadMetersPerYear,
    FOIMean,
    FOI99thPercentile,
    FOI95thPercentile,
    FOI90thPercentile,
    FOI85thPercentile,
    FOI80thPercentile,
    FOI75thPercentile,
    FOI70thPercentile,
    FOI65thPercentile,
    FOI60thPercentile,
    FOI55thPercentile,
    FOI50thPercentile,
    FOI45thPercentile,
    FOI40thPercentile,
    FOI35thPercentile,
    FOI30thPercentile,
    FOI25thPercentile,
    FOI20thPercentile,
    FOI15thPercentile,
    FOI10thPercentile,
    FOI99thPercentileMean,
    FOI95thPercentileMean,
    FOI90thPercentileMean,
    FOI85thPercentileMean,
    FOI80thPercentileMean,
    FOI75thPercentileMean,
    FOI70thPercentileMean,
    FOI65thPercentileMean,
    FOI60thPercentileMean,
    FOI55thPercentileMean,
    FOI50thPercentileMean,
    FOI45thPercentileMean,
    FOI40thPercentileMean,
    FOI35thPercentileMean,
    FOI30thPercentileMean,
    FOI25thPercentileMean,
    FOI20thPercentileMean,
    FOI15thPercentileMean,
    FOI10thPercentileMean,
    FOI99thPercentileMeanDistanceMeters,
    FOI95thPercentileMeanDistanceMeters,
    FOI90thPercentileMeanDistanceMeters,
    FOI85thPercentileMeanDistanceMeters,
    FOI80thPercentileMeanDistanceMeters,
    FOI75thPercentileMeanDistanceMeters,
    FOI70thPercentileMeanDistanceMeters,
    FOI65thPercentileMeanDistanceMeters,
    FOI60thPercentileMeanDistanceMeters,
    FOI55thPercentileMeanDistanceMeters,
    FOI50thPercentileMeanDistanceMeters,
    FOI45thPercentileMeanDistanceMeters,
    FOI40thPercentileMeanDistanceMeters,
    FOI35thPercentileMeanDistanceMeters,
    FOI30thPercentileMeanDistanceMeters,
    FOI25thPercentileMeanDistanceMeters,
    FOI20thPercentileMeanDistanceMeters,
    FOI15thPercentileMeanDistanceMeters,
    FOI10thPercentileMeanDistanceMeters,
    FOI99thPercentileSpreadMetersPerYear,
    FOI95thPercentileSpreadMetersPerYear,
    FOI90thPercentileSpreadMetersPerYear,
    FOI85thPercentileSpreadMetersPerYear,
    FOI80thPercentileSpreadMetersPerYear,
    FOI75thPercentileSpreadMetersPerYear,
    FOI70thPercentileSpreadMetersPerYear,
    FOI65thPercentileSpreadMetersPerYear,
    FOI60thPercentileSpreadMetersPerYear,
    FOI55thPercentileSpreadMetersPerYear,
    FOI50thPercentileSpreadMetersPerYear,
    FOI45thPercentileSpreadMetersPerYear,
    FOI40thPercentileSpreadMetersPerYear,
    FOI35thPercentileSpreadMetersPerYear,
    FOI30thPercentileSpreadMetersPerYear,
    FOI25thPercentileSpreadMetersPerYear,
    FOI20thPercentileSpreadMetersPerYear,
    FOI15thPercentileSpreadMetersPerYear,
    FOI10thPercentileSpreadMetersPerYear,
    TotalAnnualMortality,
    TotalHealthyBiomass,
    TotalInfectedBiomass,
    TotalIgnoredBiomass,
    TotalBiomass,
    ProportionInfectedBiomass,
    ProportionHostInfectedBiomass,
    TotalHealthyBiomassChange,
    TotalInfectedBiomassChange,
    TotalIgnoredBiomassChange,
    TotalBiomassChange,
    TotalHealthyBiomassChangePercentage,
    TotalInfectedBiomassChangePercentage,
    TotalIgnoredBiomassChangePercentage,
    TotalBiomassChangePercentage,
    TotalHealthyBiomassChangeMod,
    TotalInfectedBiomassChangeMod,
    TotalIgnoredBiomassChangeMod,
    TotalBiomassChangeMod,
}

impl Column {
    fn number(self) -> u16 {
        match self {
            Column::Timestep => 0,
            Column::InfTotal => 1,
            Column::InfMeanDistMeters => 2,
            Column::InfNew => 3,
            Column::InfectedNewChange => 4,
            Column::InfNewChangePercentage => 5,
            Column::InfNewChangeMod => 6,
            Column::InfAreaSquareMeters => 7,
            Column::InfAreaChangeSquareMeters => 8,
            Column::InfAreaChangeSquareMetersPercentage => 9,
            Column::InfectedAreaChangeSquareMetersMod => 10,
            Column::Inf99thPercentile => 11,
            Column::Inf95thPercentile => 12,
            Column::Inf90thPercentile => 13,
            Column::Inf85thPercentile => 14,
            Column::Inf80thPercentile => 15,
            Column::Inf75thPercentile => 16,
            Column::Inf70thPercentile => 17,
            Column::Inf65thPercentile => 18,
            Column::Inf60thPercentile => 19,
            Column::Inf55thPercentile => 20,
            Column::Inf50thPercentile => 21,
            Column::Inf45thPercentile => 22,
            Column::Inf40thPercentile => 23,
            Column::Inf35thPercentile => 24,
            Column::Inf30thPercentile => 25,
            Column::Inf25thPercentile => 26,
            Column::Inf20thPercentile => 27,
            Column::Inf15thPercentile => 28,
            Column::Inf10thPercentile => 29,
            Column::Inf99thPercentileMeanDistanceMeters => 30,
            Column::Inf95thPercentileMeanDistanceMeters => 31,
            Column::Inf90thPercentileMeanDistanceMeters => 32,
            Column::Inf85thPercentileMeanDistanceMeters => 33,
            Column::Inf80thPercentileMeanDistanceMeters => 34,
            Column::Inf75thPercentileMeanDistanceMeters => 35,
            Column::Inf70thPercentileMeanDistanceMeters => 36,
            Column::Inf65thPercentileMeanDistanceMeters => 37,
            Column::Inf60thPercentileMeanDistanceMeters => 38,
            Column::Inf55thPercentileMeanDistanceMeters => 39,
            Column::Inf50thPercentileMeanDistanceMeters => 40,
            Column::Inf45thPercentileMeanDistanceMeters => 41,
            Column::Inf40thPercentileMeanDistanceMeters => 42,
            Column::Inf35thPercentileMeanDistanceMeters => 43,
            Column::Inf30thPercentileMeanDistanceMeters => 44,
            Column::Inf25thPercentileMeanDistanceMeters => 45,
            Column::Inf20thPercentileMeanDistanceMeters => 46,
            Column::Inf15thPercentileMeanDistanceMeters => 47,
            Column::Inf10thPercentileMeanDistanceMeters => 48,
            Column::Inf99thPercentileSpreadMetersPerYear => 49,
            Column::Inf95thPercentileSpreadMetersPerYear => 50,
            Column::Inf90thPercentileSpreadMetersPerYear => 51,
            Column::Inf85thPercentileSpreadMetersPerYear => 52,
            Column::Inf80thPercentileSpreadMetersPerYear => 53,
            Column::Inf75thPercentileSpreadMetersPerYear => 54,
            Column::Inf70thPercentileSpreadMetersPerYear => 55,
            Column::Inf65thPercentileSpreadMetersPerYear => 56,
            Column::Inf60thPercentileSpreadMetersPerYear => 57,
            Column::Inf55thPercentileSpreadMetersPerYear => 58,
            Column::Inf50thPercentileSpreadMetersPerYear => 59,
            Column::Inf45thPercentileSpreadMetersPerYear => 60,
            Column::Inf40thPercentileSpreadMetersPerYear => 61,
            Column::Inf35thPercentileSpreadMetersPerYear => 62,
            Column::Inf30thPercentileSpreadMetersPerYear => 63,
            Column::Inf25thPercentileSpreadMetersPerYear => 64,
            Column::Inf20thPercentileSpreadMetersPerYear => 65,
            Column::Inf15thPercentileSpreadMetersPerYear => 66,
            Column::Inf10thPercentileSpreadMetersPerYear => 67,
            Column::FOIMean => 68,
            Column::FOI99thPercentile => 69,
            Column::FOI95thPercentile => 70,
            Column::FOI90thPercentile => 71,
            Column::FOI85thPercentile => 72,
            Column::FOI80thPercentile => 73,
            Column::FOI75thPercentile => 74,
            Column::FOI70thPercentile => 75,
            Column::FOI65thPercentile => 76,
            Column::FOI60thPercentile => 77,
            Column::FOI55thPercentile => 78,
            Column::FOI50thPercentile => 79,
            Column::FOI45thPercentile => 80,
            Column::FOI40thPercentile => 81,
            Column::FOI35thPercentile => 82,
            Column::FOI30thPercentile => 83,
            Column::FOI25thPercentile => 84,
            Column::FOI20thPercentile => 85,
            Column::FOI15thPercentile => 86,
            Column::FOI10thPercentile => 87,
            Column::FOI99thPercentileMean => 88,
            Column::FOI95thPercentileMean => 89,
            Column::FOI90thPercentileMean => 90,
            Column::FOI85thPercentileMean => 91,
            Column::FOI80thPercentileMean => 92,
            Column::FOI75thPercentileMean => 93,
            Column::FOI70thPercentileMean => 94,
            Column::FOI65thPercentileMean => 95,
            Column::FOI60thPercentileMean => 96,
            Column::FOI55thPercentileMean => 97,
            Column::FOI50thPercentileMean => 98,
            Column::FOI45thPercentileMean => 99,
            Column::FOI40thPercentileMean => 100,
            Column::FOI35thPercentileMean => 101,
            Column::FOI30thPercentileMean => 102,
            Column::FOI25thPercentileMean => 103,
            Column::FOI20thPercentileMean => 104,
            Column::FOI15thPercentileMean => 105,
            Column::FOI10thPercentileMean => 106,
            Column::FOI99thPercentileMeanDistanceMeters => 107,
            Column::FOI95thPercentileMeanDistanceMeters => 108,
            Column::FOI90thPercentileMeanDistanceMeters => 109,
            Column::FOI85thPercentileMeanDistanceMeters => 110,
            Column::FOI80thPercentileMeanDistanceMeters => 111,
            Column::FOI75thPercentileMeanDistanceMeters => 112,
            Column::FOI70thPercentileMeanDistanceMeters => 113,
            Column::FOI65thPercentileMeanDistanceMeters => 114,
            Column::FOI60thPercentileMeanDistanceMeters => 115,
            Column::FOI55thPercentileMeanDistanceMeters => 116,
            Column::FOI50thPercentileMeanDistanceMeters => 117,
            Column::FOI45thPercentileMeanDistanceMeters => 118,
            Column::FOI40thPercentileMeanDistanceMeters => 119,
            Column::FOI35thPercentileMeanDistanceMeters => 120,
            Column::FOI30thPercentileMeanDistanceMeters => 121,
            Column::FOI25thPercentileMeanDistanceMeters => 122,
            Column::FOI20thPercentileMeanDistanceMeters => 123,
            Column::FOI15thPercentileMeanDistanceMeters => 124,
            Column::FOI10thPercentileMeanDistanceMeters => 125,
            Column::FOI99thPercentileSpreadMetersPerYear => 126,
            Column::FOI95thPercentileSpreadMetersPerYear => 127,
            Column::FOI90thPercentileSpreadMetersPerYear => 128,
            Column::FOI85thPercentileSpreadMetersPerYear => 129,
            Column::FOI80thPercentileSpreadMetersPerYear => 130,
            Column::FOI75thPercentileSpreadMetersPerYear => 131,
            Column::FOI70thPercentileSpreadMetersPerYear => 132,
            Column::FOI65thPercentileSpreadMetersPerYear => 133,
            Column::FOI60thPercentileSpreadMetersPerYear => 134,
            Column::FOI55thPercentileSpreadMetersPerYear => 135,
            Column::FOI50thPercentileSpreadMetersPerYear => 136,
            Column::FOI45thPercentileSpreadMetersPerYear => 137,
            Column::FOI40thPercentileSpreadMetersPerYear => 138,
            Column::FOI35thPercentileSpreadMetersPerYear => 139,
            Column::FOI30thPercentileSpreadMetersPerYear => 140,
            Column::FOI25thPercentileSpreadMetersPerYear => 141,
            Column::FOI20thPercentileSpreadMetersPerYear => 142,
            Column::FOI15thPercentileSpreadMetersPerYear => 143,
            Column::FOI10thPercentileSpreadMetersPerYear => 144,
            Column::TotalAnnualMortality => 145,
            Column::TotalHealthyBiomass => 146,
            Column::TotalInfectedBiomass => 147,
            Column::TotalIgnoredBiomass => 148,
            Column::TotalBiomass => 149,
            Column::ProportionInfectedBiomass => 150,
            Column::ProportionHostInfectedBiomass => 151,
            Column::TotalHealthyBiomassChange => 152,
            Column::TotalInfectedBiomassChange => 153,
            Column::TotalIgnoredBiomassChange => 154,
            Column::TotalBiomassChange => 155,
            Column::TotalHealthyBiomassChangePercentage => 156,
            Column::TotalInfectedBiomassChangePercentage => 157,
            Column::TotalIgnoredBiomassChangePercentage => 158,
            Column::TotalBiomassChangePercentage => 159,
            Column::TotalHealthyBiomassChangeMod => 160,
            Column::TotalInfectedBiomassChangeMod => 161,
            Column::TotalIgnoredBiomassChangeMod => 162,
            Column::TotalBiomassChangeMod => 163,
        }
    }

    fn to_string(self) -> String {
        match self {
            Column::Timestep => "timestep",
            Column::InfTotal => "inf_total",
            Column::InfMeanDistMeters => "inf_mean_dist_m",
            Column::InfNew => "inf_new",
            Column::InfectedNewChange => "inf_new_change",
            Column::InfNewChangePercentage => "inf_new_change_percent",
            Column::InfNewChangeMod => "inf_new_change_mod",
            Column::InfAreaSquareMeters => "inf_area_m2",
            Column::InfAreaChangeSquareMeters => "inf_area_change_m2",
            Column::InfAreaChangeSquareMetersPercentage => "inf_area_change_m2_percent",
            Column::InfectedAreaChangeSquareMetersMod => "inf_area_change_m2_mod",
            Column::Inf99thPercentile => "inf_99th_p",
            Column::Inf95thPercentile => "inf_95th_p",
            Column::Inf90thPercentile => "inf_90th_p",
            Column::Inf85thPercentile => "inf_85th_p",
            Column::Inf80thPercentile => "inf_80th_p",
            Column::Inf75thPercentile => "inf_75th_p",
            Column::Inf70thPercentile => "inf_70th_p",
            Column::Inf65thPercentile => "inf_65th_p",
            Column::Inf60thPercentile => "inf_60th_p",
            Column::Inf55thPercentile => "inf_55th_p",
            Column::Inf50thPercentile => "inf_50th_p",
            Column::Inf45thPercentile => "inf_45th_p",
            Column::Inf40thPercentile => "inf_40th_p",
            Column::Inf35thPercentile => "inf_35th_p",
            Column::Inf30thPercentile => "inf_30th_p",
            Column::Inf25thPercentile => "inf_25th_p",
            Column::Inf20thPercentile => "inf_20th_p",
            Column::Inf15thPercentile => "inf_15th_p",
            Column::Inf10thPercentile => "inf_10th_p",
            Column::Inf99thPercentileMeanDistanceMeters => "inf_99th_p_mean_dist_m",
            Column::Inf95thPercentileMeanDistanceMeters => "inf_95th_p_mean_dist_m",
            Column::Inf90thPercentileMeanDistanceMeters => "inf_90th_p_mean_dist_m",
            Column::Inf85thPercentileMeanDistanceMeters => "inf_85th_p_mean_dist_m",
            Column::Inf80thPercentileMeanDistanceMeters => "inf_80th_p_mean_dist_m",
            Column::Inf75thPercentileMeanDistanceMeters => "inf_75th_p_mean_dist_m",
            Column::Inf70thPercentileMeanDistanceMeters => "inf_70th_p_mean_dist_m",
            Column::Inf65thPercentileMeanDistanceMeters => "inf_65th_p_mean_dist_m",
            Column::Inf60thPercentileMeanDistanceMeters => "inf_60th_p_mean_dist_m",
            Column::Inf55thPercentileMeanDistanceMeters => "inf_55th_p_mean_dist_m",
            Column::Inf50thPercentileMeanDistanceMeters => "inf_50th_p_mean_dist_m",
            Column::Inf45thPercentileMeanDistanceMeters => "inf_45th_p_mean_dist_m",
            Column::Inf40thPercentileMeanDistanceMeters => "inf_40th_p_mean_dist_m",
            Column::Inf35thPercentileMeanDistanceMeters => "inf_35th_p_mean_dist_m",
            Column::Inf30thPercentileMeanDistanceMeters => "inf_30th_p_mean_dist_m",
            Column::Inf25thPercentileMeanDistanceMeters => "inf_25th_p_mean_dist_m",
            Column::Inf20thPercentileMeanDistanceMeters => "inf_20th_p_mean_dist_m",
            Column::Inf15thPercentileMeanDistanceMeters => "inf_15th_p_mean_dist_m",
            Column::Inf10thPercentileMeanDistanceMeters => "inf_10th_p_mean_dist_m",
            Column::Inf99thPercentileSpreadMetersPerYear => "inf_99th_p_spread_mpy",
            Column::Inf95thPercentileSpreadMetersPerYear => "inf_95th_p_spread_mpy",
            Column::Inf90thPercentileSpreadMetersPerYear => "inf_90th_p_spread_mpy",
            Column::Inf85thPercentileSpreadMetersPerYear => "inf_85th_p_spread_mpy",
            Column::Inf80thPercentileSpreadMetersPerYear => "inf_80th_p_spread_mpy",
            Column::Inf75thPercentileSpreadMetersPerYear => "inf_75th_p_spread_mpy",
            Column::Inf70thPercentileSpreadMetersPerYear => "inf_70th_p_spread_mpy",
            Column::Inf65thPercentileSpreadMetersPerYear => "inf_65th_p_spread_mpy",
            Column::Inf60thPercentileSpreadMetersPerYear => "inf_60th_p_spread_mpy",
            Column::Inf55thPercentileSpreadMetersPerYear => "inf_55th_p_spread_mpy",
            Column::Inf50thPercentileSpreadMetersPerYear => "inf_50th_p_spread_mpy",
            Column::Inf45thPercentileSpreadMetersPerYear => "inf_45th_p_spread_mpy",
            Column::Inf40thPercentileSpreadMetersPerYear => "inf_40th_p_spread_mpy",
            Column::Inf35thPercentileSpreadMetersPerYear => "inf_35th_p_spread_mpy",
            Column::Inf30thPercentileSpreadMetersPerYear => "inf_30th_p_spread_mpy",
            Column::Inf25thPercentileSpreadMetersPerYear => "inf_25th_p_spread_mpy",
            Column::Inf20thPercentileSpreadMetersPerYear => "inf_20th_p_spread_mpy",
            Column::Inf15thPercentileSpreadMetersPerYear => "inf_15th_p_spread_mpy",
            Column::Inf10thPercentileSpreadMetersPerYear => "inf_10th_p_spread_mpy",
            Column::FOIMean => "foi_mean",
            Column::FOI99thPercentile => "foi_99th_p",
            Column::FOI95thPercentile => "foi_95th_p",
            Column::FOI90thPercentile => "foi_90th_p",
            Column::FOI85thPercentile => "foi_85th_p",
            Column::FOI80thPercentile => "foi_80th_p",
            Column::FOI75thPercentile => "foi_75th_p",
            Column::FOI70thPercentile => "foi_70th_p",
            Column::FOI65thPercentile => "foi_65th_p",
            Column::FOI60thPercentile => "foi_60th_p",
            Column::FOI55thPercentile => "foi_55th_p",
            Column::FOI50thPercentile => "foi_50th_p",
            Column::FOI45thPercentile => "foi_45th_p",
            Column::FOI40thPercentile => "foi_40th_p",
            Column::FOI35thPercentile => "foi_35th_p",
            Column::FOI30thPercentile => "foi_30th_p",
            Column::FOI25thPercentile => "foi_25th_p",
            Column::FOI20thPercentile => "foi_20th_p",
            Column::FOI15thPercentile => "foi_15th_p",
            Column::FOI10thPercentile => "foi_10th_p",
            Column::FOI99thPercentileMean => "foi_99th_p_mean",
            Column::FOI95thPercentileMean => "foi_95th_p_mean",
            Column::FOI90thPercentileMean => "foi_90th_p_mean",
            Column::FOI85thPercentileMean => "foi_85th_p_mean",
            Column::FOI80thPercentileMean => "foi_80th_p_mean",
            Column::FOI75thPercentileMean => "foi_75th_p_mean",
            Column::FOI70thPercentileMean => "foi_70th_p_mean",
            Column::FOI65thPercentileMean => "foi_65th_p_mean",
            Column::FOI60thPercentileMean => "foi_60th_p_mean",
            Column::FOI55thPercentileMean => "foi_55th_p_mean",
            Column::FOI50thPercentileMean => "foi_50th_p_mean",
            Column::FOI45thPercentileMean => "foi_45th_p_mean",
            Column::FOI40thPercentileMean => "foi_40th_p_mean",
            Column::FOI35thPercentileMean => "foi_35th_p_mean",
            Column::FOI30thPercentileMean => "foi_30th_p_mean",
            Column::FOI25thPercentileMean => "foi_25th_p_mean",
            Column::FOI20thPercentileMean => "foi_20th_p_mean",
            Column::FOI15thPercentileMean => "foi_15th_p_mean",
            Column::FOI10thPercentileMean => "foi_10th_p_mean",
            Column::FOI99thPercentileMeanDistanceMeters => "foi_99th_p_mean_dist_m",
            Column::FOI95thPercentileMeanDistanceMeters => "foi_95th_p_mean_dist_m",
            Column::FOI90thPercentileMeanDistanceMeters => "foi_90th_p_mean_dist_m",
            Column::FOI85thPercentileMeanDistanceMeters => "foi_85th_p_mean_dist_m",
            Column::FOI80thPercentileMeanDistanceMeters => "foi_80th_p_mean_dist_m",
            Column::FOI75thPercentileMeanDistanceMeters => "foi_75th_p_mean_dist_m",
            Column::FOI70thPercentileMeanDistanceMeters => "foi_70th_p_mean_dist_m",
            Column::FOI65thPercentileMeanDistanceMeters => "foi_65th_p_mean_dist_m",
            Column::FOI60thPercentileMeanDistanceMeters => "foi_60th_p_mean_dist_m",
            Column::FOI55thPercentileMeanDistanceMeters => "foi_55th_p_mean_dist_m",
            Column::FOI50thPercentileMeanDistanceMeters => "foi_50th_p_mean_dist_m",
            Column::FOI45thPercentileMeanDistanceMeters => "foi_45th_p_mean_dist_m",
            Column::FOI40thPercentileMeanDistanceMeters => "foi_40th_p_mean_dist_m",
            Column::FOI35thPercentileMeanDistanceMeters => "foi_35th_p_mean_dist_m",
            Column::FOI30thPercentileMeanDistanceMeters => "foi_30th_p_mean_dist_m",
            Column::FOI25thPercentileMeanDistanceMeters => "foi_25th_p_mean_dist_m",
            Column::FOI20thPercentileMeanDistanceMeters => "foi_20th_p_mean_dist_m",
            Column::FOI15thPercentileMeanDistanceMeters => "foi_15th_p_mean_dist_m",
            Column::FOI10thPercentileMeanDistanceMeters => "foi_10th_p_mean_dist_m",
            Column::FOI99thPercentileSpreadMetersPerYear => "foi_99th_p_spread_mpy",
            Column::FOI95thPercentileSpreadMetersPerYear => "foi_95th_p_spread_mpy",
            Column::FOI90thPercentileSpreadMetersPerYear => "foi_90th_p_spread_mpy",
            Column::FOI85thPercentileSpreadMetersPerYear => "foi_85th_p_spread_mpy",
            Column::FOI80thPercentileSpreadMetersPerYear => "foi_80th_p_spread_mpy",
            Column::FOI75thPercentileSpreadMetersPerYear => "foi_75th_p_spread_mpy",
            Column::FOI70thPercentileSpreadMetersPerYear => "foi_70th_p_spread_mpy",
            Column::FOI65thPercentileSpreadMetersPerYear => "foi_65th_p_spread_mpy",
            Column::FOI60thPercentileSpreadMetersPerYear => "foi_60th_p_spread_mpy",
            Column::FOI55thPercentileSpreadMetersPerYear => "foi_55th_p_spread_mpy",
            Column::FOI50thPercentileSpreadMetersPerYear => "foi_50th_p_spread_mpy",
            Column::FOI45thPercentileSpreadMetersPerYear => "foi_45th_p_spread_mpy",
            Column::FOI40thPercentileSpreadMetersPerYear => "foi_40th_p_spread_mpy",
            Column::FOI35thPercentileSpreadMetersPerYear => "foi_35th_p_spread_mpy",
            Column::FOI30thPercentileSpreadMetersPerYear => "foi_30th_p_spread_mpy",
            Column::FOI25thPercentileSpreadMetersPerYear => "foi_25th_p_spread_mpy",
            Column::FOI20thPercentileSpreadMetersPerYear => "foi_20th_p_spread_mpy",
            Column::FOI15thPercentileSpreadMetersPerYear => "foi_15th_p_spread_mpy",
            Column::FOI10thPercentileSpreadMetersPerYear => "foi_10th_p_spread_mpy",
            Column::TotalAnnualMortality => "t_annual_mortality",
            Column::TotalHealthyBiomass => "t_hea_biomass",
            Column::TotalInfectedBiomass => "t_inf_biomass",
            Column::TotalIgnoredBiomass => "t_ign_biomass",
            Column::TotalBiomass => "t_biomass",
            Column::ProportionInfectedBiomass => "prop_inf_biomass",
            Column::ProportionHostInfectedBiomass => "prop_host_inf_biomass",
            Column::TotalHealthyBiomassChange => "t_hea_biomass_change",
            Column::TotalInfectedBiomassChange => "t_inf_biomass_change",
            Column::TotalIgnoredBiomassChange => "t_ign_biomass_change",
            Column::TotalBiomassChange => "t_biomass_change",
            Column::TotalHealthyBiomassChangePercentage => "t_hea_biomass_change_percent",
            Column::TotalInfectedBiomassChangePercentage => "t_inf_biomass_change_percent",
            Column::TotalIgnoredBiomassChangePercentage => "t_ign_biomass_change_percent",
            Column::TotalBiomassChangePercentage => "t_biomass_change_percent",
            Column::TotalHealthyBiomassChangeMod => "t_hea_biomass_change_mod",
            Column::TotalInfectedBiomassChangeMod => "t_inf_biomass_change_mod",
            Column::TotalIgnoredBiomassChangeMod => "t_ign_biomass_change_mod",
            Column::TotalBiomassChangeMod => "t_biomass_change_mod",
        }
        .to_string()
    }

    fn description(self) -> String {
        match self {
            Column::Timestep => "timestep",
            Column::InfTotal => "total number of infected sites",
            Column::InfMeanDistMeters => "infection mean distance from source in meters",
            Column::InfNew => "newly infected sites",
            Column::InfectedNewChange => "newly infected sites change",
            Column::InfNewChangePercentage => "newly infection sites percent",
            Column::InfNewChangeMod => "newly infected sites change modifier",
            Column::InfAreaSquareMeters => "infected area square meters",
            Column::InfAreaChangeSquareMeters => "infected area change square meters",
            Column::InfAreaChangeSquareMetersPercentage => {
                "infected area change square meters percent"
            }
            Column::InfectedAreaChangeSquareMetersMod => "infection area change modifier",
            Column::Inf99thPercentile => "infection 99th percentile",
            Column::Inf95thPercentile => "infection 95th percentile",
            Column::Inf90thPercentile => "infection 90th percentile",
            Column::Inf85thPercentile => "infection 85th percentile",
            Column::Inf80thPercentile => "infection 80th percentile",
            Column::Inf75thPercentile => "infection 75th percentile",
            Column::Inf70thPercentile => "infection 70th percentile",
            Column::Inf65thPercentile => "infection 65th percentile",
            Column::Inf60thPercentile => "infection 60th percentile",
            Column::Inf55thPercentile => "infection 55th percentile",
            Column::Inf50thPercentile => "infection 50th percentile",
            Column::Inf45thPercentile => "infection 45th percentile",
            Column::Inf40thPercentile => "infection 40th percentile",
            Column::Inf35thPercentile => "infection 35th percentile",
            Column::Inf30thPercentile => "infection 30th percentile",
            Column::Inf25thPercentile => "infection 25th percentile",
            Column::Inf20thPercentile => "infection 20th percentile",
            Column::Inf15thPercentile => "infection 15th percentile",
            Column::Inf10thPercentile => "infection 10th percentile",
            Column::Inf99thPercentileMeanDistanceMeters => {
                "infection mean distance of 99th percentile from source"
            }
            Column::Inf95thPercentileMeanDistanceMeters => {
                "infection mean distance of 95th percentile from source"
            }
            Column::Inf99thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 99th percentile"
            }
            Column::Inf95thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 95th percentile"
            }
            Column::FOIMean => "force of infection mean",
            Column::FOI99thPercentile => "force of infection 99th percentile",
            Column::FOI95thPercentile => "force of infection 95th percentile",
            Column::FOI90thPercentile => "force of infection 90th percentile",
            Column::FOI85thPercentile => "force of infection 85th percentile",
            Column::FOI80thPercentile => "force of infection 80th percentile",
            Column::FOI75thPercentile => "force of infection 75th percentile",
            Column::FOI70thPercentile => "force of infection 70th percentile",
            Column::FOI65thPercentile => "force of infection 65th percentile",
            Column::FOI60thPercentile => "force of infection 60th percentile",
            Column::FOI55thPercentile => "force of infection 55th percentile",
            Column::FOI50thPercentile => "force of infection 50th percentile",
            Column::FOI45thPercentile => "force of infection 45th percentile",
            Column::FOI40thPercentile => "force of infection 40th percentile",
            Column::FOI35thPercentile => "force of infection 35th percentile",
            Column::FOI30thPercentile => "force of infection 30th percentile",
            Column::FOI25thPercentile => "force of infection 25th percentile",
            Column::FOI20thPercentile => "force of infection 20th percentile",
            Column::FOI15thPercentile => "force of infection 15th percentile",
            Column::FOI10thPercentile => "force of infection 10th percentile",
            Column::FOI99thPercentileMean => "force of infection 99th percentile mean",
            Column::FOI95thPercentileMean => "force of infection 95th percentile mean",
            Column::FOI90thPercentileMean => "force of infection 90th percentile mean",
            Column::FOI85thPercentileMean => "force of infection 85th percentile mean",
            Column::FOI80thPercentileMean => "force of infection 80th percentile mean",
            Column::FOI75thPercentileMean => "force of infection 75th percentile mean",
            Column::FOI70thPercentileMean => "force of infection 70th percentile mean",
            Column::FOI65thPercentileMean => "force of infection 65th percentile mean",
            Column::FOI60thPercentileMean => "force of infection 60th percentile mean",
            Column::FOI55thPercentileMean => "force of infection 55th percentile mean",
            Column::FOI50thPercentileMean => "force of infection 50th percentile mean",
            Column::FOI45thPercentileMean => "force of infection 45th percentile mean",
            Column::FOI40thPercentileMean => "force of infection 40th percentile mean",
            Column::FOI35thPercentileMean => "force of infection 35th percentile mean",
            Column::FOI30thPercentileMean => "force of infection 30th percentile mean",
            Column::FOI25thPercentileMean => "force of infection 25th percentile mean",
            Column::FOI20thPercentileMean => "force of infection 20th percentile mean",
            Column::FOI15thPercentileMean => "force of infection 15th percentile mean",
            Column::FOI10thPercentileMean => "force of infection 10th percentile mean",
            Column::FOI99thPercentileMeanDistanceMeters => {
                "force of infection 99th percentile mean distance from source in meters"
            }
            Column::FOI95thPercentileMeanDistanceMeters => {
                "force of infection 95th percentile mean distance from source in meters"
            }
            Column::FOI90thPercentileMeanDistanceMeters => {
                "force of infection 90th percentile mean distance from source in meters"
            }
            Column::FOI85thPercentileMeanDistanceMeters => {
                "force of infection 85th percentile mean distance from source in meters"
            }
            Column::FOI80thPercentileMeanDistanceMeters => {
                "force of infection 80th percentile mean distance from source in meters"
            }
            Column::FOI75thPercentileMeanDistanceMeters => {
                "force of infection 75th percentile mean distance from source in meters"
            }
            Column::FOI70thPercentileMeanDistanceMeters => {
                "force of infection 70th percentile mean distance from source in meters"
            }
            Column::FOI65thPercentileMeanDistanceMeters => {
                "force of infection 65th percentile mean distance from source in meters"
            }
            Column::FOI60thPercentileMeanDistanceMeters => {
                "force of infection 60th percentile mean distance from source in meters"
            }
            Column::FOI55thPercentileMeanDistanceMeters => {
                "force of infection 55th percentile mean distance from source in meters"
            }
            Column::FOI50thPercentileMeanDistanceMeters => {
                "force of infection 50th percentile mean distance from source in meters"
            }
            Column::FOI45thPercentileMeanDistanceMeters => {
                "force of infection 45th percentile mean distance from source in meters"
            }
            Column::FOI40thPercentileMeanDistanceMeters => {
                "force of infection 40th percentile mean distance from source in meters"
            }
            Column::FOI35thPercentileMeanDistanceMeters => {
                "force of infection 35th percentile mean distance from source in meters"
            }
            Column::FOI30thPercentileMeanDistanceMeters => {
                "force of infection 30th percentile mean distance from source in meters"
            }
            Column::FOI25thPercentileMeanDistanceMeters => {
                "force of infection 25th percentile mean distance from source in meters"
            }
            Column::FOI20thPercentileMeanDistanceMeters => {
                "force of infection 20th percentile mean distance from source in meters"
            }
            Column::FOI15thPercentileMeanDistanceMeters => {
                "force of infection 15th percentile mean distance from source in meters"
            }
            Column::FOI10thPercentileMeanDistanceMeters => {
                "force of infection 10th percentile mean distance from source in meters"
            }
            Column::TotalAnnualMortality => "annual mortality (conditional formatting indicates the min (green), max (red) and median (yellow) values)",
            Column::TotalHealthyBiomass => "total healthy biomass",
            Column::TotalInfectedBiomass => "total infected biomass",
            Column::TotalIgnoredBiomass => "total ignored biomass",
            Column::TotalBiomass => "total biomass",
            Column::ProportionInfectedBiomass => {
                "proportion of infected biomass (between 0.0 and 1.0)"
            }
            Column::ProportionHostInfectedBiomass => {
                "proportion of host infected biomass (between 0.0 and 1.0)"
            }
            Column::TotalHealthyBiomassChange => "total healthy biomass change",
            Column::TotalInfectedBiomassChange => "total infected biomass change",
            Column::TotalIgnoredBiomassChange => "total ignored biomass change",
            Column::TotalBiomassChange => "total biomass change",
            Column::TotalHealthyBiomassChangePercentage => "total healthy biomass change percent",
            Column::TotalInfectedBiomassChangePercentage => "total infected biomass change percent",
            Column::TotalIgnoredBiomassChangePercentage => "total ignored biomass change percent",
            Column::TotalBiomassChangePercentage => "total biomass change percent",
            Column::TotalHealthyBiomassChangeMod => "total healthy biomass change modifier",
            Column::TotalInfectedBiomassChangeMod => "total infected biomass change modifier",
            Column::TotalIgnoredBiomassChangeMod => "total ignored biomass change modifier",
            Column::TotalBiomassChangeMod => "total biomass change modifier",
            Column::Inf90thPercentileMeanDistanceMeters => {
                "infection mean distance of 90th percentile from source"
            }
            Column::Inf85thPercentileMeanDistanceMeters => {
                "infection mean distance of 85th percentile from source"
            }
            Column::Inf80thPercentileMeanDistanceMeters => {
                "infection mean distance of 80th percentile from source"
            }
            Column::Inf75thPercentileMeanDistanceMeters => {
                "infection mean distance of 75th percentile from source"
            }
            Column::Inf70thPercentileMeanDistanceMeters => {
                "infection mean distance of 70th percentile from source"
            }
            Column::Inf65thPercentileMeanDistanceMeters => {
                "infection mean distance of 65th percentile from source"
            }
            Column::Inf60thPercentileMeanDistanceMeters => {
                "infection mean distance of 60th percentile from source"
            }
            Column::Inf55thPercentileMeanDistanceMeters => {
                "infection mean distance of 55th percentile from source"
            }
            Column::Inf50thPercentileMeanDistanceMeters => {
                "infection mean distance of 50th percentile from source"
            }
            Column::Inf45thPercentileMeanDistanceMeters => {
                "infection mean distance of 45th percentile from source"
            }
            Column::Inf40thPercentileMeanDistanceMeters => {
                "infection mean distance of 40th percentile from source"
            }
            Column::Inf35thPercentileMeanDistanceMeters => {
                "infection mean distance of 35th percentile from source"
            }
            Column::Inf30thPercentileMeanDistanceMeters => {
                "infection mean distance of 30th percentile from source"
            }
            Column::Inf25thPercentileMeanDistanceMeters => {
                "infection mean distance of 25th percentile from source"
            }
            Column::Inf20thPercentileMeanDistanceMeters => {
                "infection mean distance of 20th percentile from source"
            }
            Column::Inf15thPercentileMeanDistanceMeters => {
                "infection mean distance of 15th percentile from source"
            }
            Column::Inf10thPercentileMeanDistanceMeters => {
                "infection mean distance of 10th percentile from source"
            }
            Column::Inf90thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 90th percentile"
            }
            Column::Inf85thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 85th percentile"
            }
            Column::Inf80thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 80th percentile"
            }
            Column::Inf75thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 75th percentile"
            }
            Column::Inf70thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 70th percentile"
            }
            Column::Inf65thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 65th percentile"
            }
            Column::Inf60thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 60th percentile"
            }
            Column::Inf55thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 55th percentile"
            }
            Column::Inf50thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 50th percentile"
            }
            Column::Inf45thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 45th percentile"
            }
            Column::Inf40thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 40th percentile"
            }
            Column::Inf35thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 35th percentile"
            }
            Column::Inf30thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 30th percentile"
            }
            Column::Inf25thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 25th percentile"
            }
            Column::Inf20thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 20th percentile"
            }
            Column::Inf15thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 15th percentile"
            }
            Column::Inf10thPercentileSpreadMetersPerYear => {
                "infection spread rate in meters per year from 10th percentile"
            }
            Column::FOI99thPercentileSpreadMetersPerYear => {
                "force of infection 99th percentile spread rate in meters per year"
            }
            Column::FOI95thPercentileSpreadMetersPerYear => {
                "force of infection 95th percentile spread rate in meters per year"
            }
            Column::FOI90thPercentileSpreadMetersPerYear => {
                "force of infection 90th percentile spread rate in meters per year"
            }
            Column::FOI85thPercentileSpreadMetersPerYear => {
                "force of infection 85th percentile spread rate in meters per year"
            }
            Column::FOI80thPercentileSpreadMetersPerYear => {
                "force of infection 80th percentile spread rate in meters per year"
            }
            Column::FOI75thPercentileSpreadMetersPerYear => {
                "force of infection 75th percentile spread rate in meters per year"
            }
            Column::FOI70thPercentileSpreadMetersPerYear => {
                "force of infection 70th percentile spread rate in meters per year"
            }
            Column::FOI65thPercentileSpreadMetersPerYear => {
                "force of infection 65th percentile spread rate in meters per year"
            }
            Column::FOI60thPercentileSpreadMetersPerYear => {
                "force of infection 60th percentile spread rate in meters per year"
            }
            Column::FOI55thPercentileSpreadMetersPerYear => {
                "force of infection 55th percentile spread rate in meters per year"
            }
            Column::FOI50thPercentileSpreadMetersPerYear => {
                "force of infection 50th percentile spread rate in meters per year"
            }
            Column::FOI45thPercentileSpreadMetersPerYear => {
                "force of infection 45th percentile spread rate in meters per year"
            }
            Column::FOI40thPercentileSpreadMetersPerYear => {
                "force of infection 40th percentile spread rate in meters per year"
            }
            Column::FOI35thPercentileSpreadMetersPerYear => {
                "force of infection 35th percentile spread rate in meters per year"
            }
            Column::FOI30thPercentileSpreadMetersPerYear => {
                "force of infection 30th percentile spread rate in meters per year"
            }
            Column::FOI25thPercentileSpreadMetersPerYear => {
                "force of infection 25th percentile spread rate in meters per year"
            }
            Column::FOI20thPercentileSpreadMetersPerYear => {
                "force of infection 20th percentile spread rate in meters per year"
            }
            Column::FOI15thPercentileSpreadMetersPerYear => {
                "force of infection 15th percentile spread rate in meters per year"
            }
            Column::FOI10thPercentileSpreadMetersPerYear => {
                "force of infection 10th percentile spread rate in meters per year"
            }
        }
        .to_string()
    }

    fn conditional_formatter(self) -> Option<ConditionalFormatter> {
        Some(match self {
            Column::InfectedNewChange => ConditionalFormatter::ChangeAbsolute,
            Column::InfNewChangePercentage => ConditionalFormatter::ChangePercentage,
            Column::InfNewChangeMod => ConditionalFormatter::ChangeModifier,
            Column::InfAreaChangeSquareMeters => ConditionalFormatter::ChangeAbsolute,
            Column::InfAreaChangeSquareMetersPercentage => ConditionalFormatter::ChangePercentage,
            Column::InfectedAreaChangeSquareMetersMod => ConditionalFormatter::ChangeModifier,
            Column::TotalHealthyBiomassChange => ConditionalFormatter::ChangeAbsolute,
            Column::TotalInfectedBiomassChange => ConditionalFormatter::ChangeAbsolute,
            Column::TotalIgnoredBiomassChange => ConditionalFormatter::ChangeAbsolute,
            Column::TotalBiomassChange => ConditionalFormatter::ChangeAbsolute,
            Column::TotalHealthyBiomassChangePercentage => ConditionalFormatter::ChangePercentage,
            Column::TotalInfectedBiomassChangePercentage => ConditionalFormatter::ChangePercentage,
            Column::TotalIgnoredBiomassChangePercentage => ConditionalFormatter::ChangePercentage,
            Column::TotalBiomassChangePercentage => ConditionalFormatter::ChangePercentage,
            Column::TotalHealthyBiomassChangeMod => ConditionalFormatter::ChangeModifier,
            Column::TotalInfectedBiomassChangeMod => ConditionalFormatter::ChangeModifier,
            Column::TotalIgnoredBiomassChangeMod => ConditionalFormatter::ChangeModifier,
            Column::TotalBiomassChangeMod => ConditionalFormatter::ChangeModifier,
            _ => return None,
        })
    }
}

mod image_grid;
use crate::image_grid::{
    GridRenderConfig, render_foi_png_gray16, render_infection_state_png, render_state_map_png,
};

#[derive(Parser, Debug)]
#[command(
    name = "dp_output_analysis",
    version,
    about = "Analyze .bin representations in a directory",
    arg_required_else_help = true
)]
struct Args {
    #[arg(value_name = "dir", required = true)]
    dir: PathBuf,
    #[arg(
        short = 'o',
        long = "output-dir",
        value_name = "output-dir",
        required = true
    )]
    output_dir: PathBuf,
    #[arg(short = 'x', required = true)]
    x: usize,
    #[arg(short = 'y', required = true)]
    y: usize,
    #[arg(short = 'w', long = "width", value_name = "width", required = true)]
    width: u32,
    #[arg(short = 'h', long = "height", value_name = "height", required = true)]
    height: u32,
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

#[derive(Serialize, Deserialize)]
struct BiomassMap {
    pub timestep: u32,
    pub width: u32,
    pub height: u32,
    pub biomass: Box<[(u64, u64, u64)]>,
}

#[derive(Serialize, Deserialize)]
struct MortalityMap {
    pub timestep: u32,
    pub width: u32,
    pub height: u32,
    pub data: Box<[u32]>,
}

#[derive(Serialize, Deserialize)]
struct StateMap {
    pub timestep: u32,
    pub width: u32,
    pub height: u32,
    pub data: Box<[bool]>,
}

#[derive(Default)]
struct MapGrouping {
    foi: Option<F64Map>,
    infection: Option<InfectionStateMap>,
    biomass: Option<BiomassMap>,
    mortality: Option<MortalityMap>,
    mortality_occurred: Option<StateMap>,
    infection_occurred: Option<StateMap>,
}

struct Foi {
    data: Box<[f64]>,
}

struct Infection {
    healthy_sites: Box<[(u32, u32)]>,
    infected_sites: Box<[(u32, u32)]>,
    ignored_sites: Box<[(u32, u32)]>,
}

struct Biomass {
    data: Box<[(u64, u64, u64)]>,
}

struct Mortality {
    data: Box<[u32]>,
}

struct StateFlags {
    data: Box<[bool]>,
}

struct CombinedState {
    timestep: u32,
    foi: Option<Foi>,
    infection: Option<Infection>,
    biomass: Option<Biomass>,
    mortality: Option<Mortality>,
    mortality_occurred: Option<StateFlags>,
    infection_occurred: Option<StateFlags>,
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
    if args.width == 0 || args.height == 0 {
        return Err("Width and height must be greater than zero".into());
    }
    let width = args.width;
    let height = args.height;
    let foi_dir = args.dir.join("foi");
    let infection_dir = args.dir.join("infection");
    let biomass_dir = args.dir.join("biomass");
    let mortality_dir = args.dir.join("mortality");
    let mortality_occurred_dir = args.dir.join("mortality_occurred");
    let infection_occurred_dir = args.dir.join("infection_occurred");

    let mut foi_files: Vec<PathBuf> = if foi_dir.is_dir() {
        WalkDir::new(&foi_dir)
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
            .collect()
    } else {
        Vec::new()
    };
    foi_files.sort_by(|a, b| compare_paths_natural(a, b));

    let mut infection_files: Vec<PathBuf> = if infection_dir.is_dir() {
        WalkDir::new(&infection_dir)
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
            .collect()
    } else {
        Vec::new()
    };
    infection_files.sort_by(|a, b| compare_paths_natural(a, b));

    let mut biomass_files: Vec<PathBuf> = if biomass_dir.is_dir() {
        WalkDir::new(&biomass_dir)
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
            .collect()
    } else {
        Vec::new()
    };
    biomass_files.sort_by(|a, b| compare_paths_natural(a, b));

    let mut mortality_files: Vec<PathBuf> = if mortality_dir.is_dir() {
        WalkDir::new(&mortality_dir)
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
            .collect()
    } else {
        Vec::new()
    };
    mortality_files.sort_by(|a, b| compare_paths_natural(a, b));

    let mut mortality_occurred_files: Vec<PathBuf> = if mortality_occurred_dir.is_dir() {
        WalkDir::new(&mortality_occurred_dir)
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
            .collect()
    } else {
        Vec::new()
    };
    mortality_occurred_files.sort_by(|a, b| compare_paths_natural(a, b));

    let mut infection_occurred_files: Vec<PathBuf> = if infection_occurred_dir.is_dir() {
        WalkDir::new(&infection_occurred_dir)
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
            .collect()
    } else {
        Vec::new()
    };
    infection_occurred_files.sort_by(|a, b| compare_paths_natural(a, b));

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
    for path in biomass_files {
        let bytes = fs::read(&path)?;
        match bincode::deserialize::<BiomassMap>(&bytes) {
            Ok(map) => {
                let entry = by_timestep.entry(map.timestep).or_default();
                entry.biomass = Some(map);
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
    for path in mortality_occurred_files {
        let bytes = fs::read(&path)?;
        match bincode::deserialize::<StateMap>(&bytes) {
            Ok(map) => {
                let timestep = map.timestep;
                let entry = by_timestep.entry(timestep).or_default();
                entry.mortality_occurred = Some(map);
            }
            Err(err) => {
                eprintln!("ERR {} {}", path.display(), err);
            }
        }
    }
    for path in infection_occurred_files {
        let bytes = fs::read(&path)?;
        match bincode::deserialize::<StateMap>(&bytes) {
            Ok(map) => {
                let timestep = map.timestep;
                let entry = by_timestep.entry(timestep).or_default();
                entry.infection_occurred = Some(map);
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
            foi: None,
            infection: None,
            biomass: None,
            mortality: None,
            mortality_occurred: None,
            infection_occurred: None,
        };

        let mut some_data_seen = false;

        if let Some(foi) = group.foi {
            if foi.timestep != timestep {
                return Err(format!("Mismatched timestep for foi at ts {}", timestep).into());
            }
            if foi.width != width || foi.height != height {
                return Err(format!(
                    "Unexpected dimensions for foi at ts {}: {}x{} (expected {}x{})",
                    timestep, foi.width, foi.height, width, height
                )
                .into());
            }
            combined_state.foi = Some(Foi { data: foi.data });
            some_data_seen = true;
        }

        if let Some(infection) = group.infection {
            if infection.timestep != timestep {
                return Err(format!("Mismatched timestep for infection at ts {}", timestep).into());
            }
            if infection.width != width || infection.height != height {
                return Err(format!(
                    "Unexpected dimensions for infection at ts {}: {}x{} (expected {}x{})",
                    timestep, infection.width, infection.height, width, height
                )
                .into());
            }
            let InfectionStateMap {
                healthy_sites,
                infected_sites,
                ignored_sites,
                ..
            } = infection;

            combined_state.infection = Some(Infection {
                healthy_sites,
                infected_sites,
                ignored_sites,
            });
            some_data_seen = true;
        }

        if let Some(biomass) = group.biomass {
            if biomass.timestep != timestep {
                return Err(format!("Mismatched timestep for biomass at ts {}", timestep).into());
            }
            if biomass.width != width || biomass.height != height {
                return Err(format!(
                    "Unexpected dimensions for biomass at ts {}: {}x{} (expected {}x{})",
                    timestep, biomass.width, biomass.height, width, height
                )
                .into());
            }
            let expected_len = width as usize * height as usize;
            if biomass.biomass.len() != expected_len {
                return Err(format!(
                    "Mismatched biomass length at ts {}: {} (expected {})",
                    timestep,
                    biomass.biomass.len(),
                    expected_len
                )
                .into());
            }
            combined_state.biomass = Some(Biomass {
                data: biomass.biomass,
            });
            some_data_seen = true;
        }

        if let Some(mortality) = group.mortality {
            if mortality.timestep != timestep {
                return Err(format!("Mismatched timestep for mortality at ts {}", timestep).into());
            }
            if mortality.width != width || mortality.height != height {
                return Err(format!(
                    "Unexpected dimensions for mortality at ts {}: {}x{} (expected {}x{})",
                    timestep, mortality.width, mortality.height, width, height
                )
                .into());
            }
            combined_state.mortality = Some(Mortality {
                data: mortality.data,
            });
            some_data_seen = true;
        }

        if let Some(mortality_occurred) = group.mortality_occurred {
            if mortality_occurred.timestep != timestep {
                return Err(format!(
                    "Mismatched timestep for mortality_occurred at ts {}",
                    timestep
                )
                .into());
            }
            if mortality_occurred.width != width || mortality_occurred.height != height {
                return Err(format!(
                    "Unexpected dimensions for mortality_occurred at ts {}: {}x{} (expected {}x{})",
                    timestep, mortality_occurred.width, mortality_occurred.height, width, height
                )
                .into());
            }
            let expected_len = width as usize * height as usize;
            if mortality_occurred.data.len() != expected_len {
                return Err(format!(
                    "Mismatched mortality_occurred length at ts {}: {} (expected {})",
                    timestep,
                    mortality_occurred.data.len(),
                    expected_len
                )
                .into());
            }
            combined_state.mortality_occurred = Some(StateFlags {
                data: mortality_occurred.data,
            });
            some_data_seen = true;
        }

        if let Some(infection_occurred) = group.infection_occurred {
            if infection_occurred.timestep != timestep {
                return Err(format!(
                    "Mismatched timestep for infection_occurred at ts {}",
                    timestep
                )
                .into());
            }
            if infection_occurred.width != width || infection_occurred.height != height {
                return Err(format!(
                    "Unexpected dimensions for infection_occurred at ts {}: {}x{} (expected {}x{})",
                    timestep, infection_occurred.width, infection_occurred.height, width, height
                )
                .into());
            }
            let expected_len = width as usize * height as usize;
            if infection_occurred.data.len() != expected_len {
                return Err(format!(
                    "Mismatched infection_occurred length at ts {}: {} (expected {})",
                    timestep,
                    infection_occurred.data.len(),
                    expected_len
                )
                .into());
            }
            combined_state.infection_occurred = Some(StateFlags {
                data: infection_occurred.data,
            });
            some_data_seen = true;
        }

        if !some_data_seen {
            eprintln!("No data seen for timestep {}", timestep);
            continue;
        }

        combined.insert(timestep, combined_state);
    }
    {
        let (x, y, cd) = (args.x, args.y, args.cd);
        let mut workbook = Workbook::new();
        {
            let worksheet = workbook.get_worksheet();
            for column in Column::iter() {
                worksheet.write_string(0, column.number(), &column.to_string())?;
                worksheet.insert_note(0, column.number(), &Note::new(&column.description()))?;
            }
        }

        let mut previous_number_of_infected_sites: usize = 0;
        let mut previous_infected_area: f64 = 0.0;
        let mut previous_newly_infected_sites: usize = 0;
        let mut previous_infection_99th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_95th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_90th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_85th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_80th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_75th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_70th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_65th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_60th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_55th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_50th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_45th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_40th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_35th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_30th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_25th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_20th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_15th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_infection_10th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_99th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_95th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_90th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_85th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_80th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_75th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_70th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_65th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_60th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_55th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_50th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_45th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_40th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_35th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_30th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_25th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_20th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_15th_percentile_mean_distance_from_source: f64 = 0.0;
        let mut previous_foi_10th_percentile_mean_distance_from_source: f64 = 0.0;
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

        let infection_source: (usize, usize) = (x, y);

        let image_cfg = GridRenderConfig::default();

        let distances_from_infection_source_1_indexed = {
            let mut map: HashMap<(usize, usize), f64> = HashMap::new();
            for i in 1..=(width as usize) {
                for j in 1..=(height as usize) {
                    let site = (i, j);
                    let distance = euclidean_distance(&infection_source, &site) * cd;
                    //exclude initial infection point
                    if distance > 0.0 {
                        map.insert(site, distance);
                    }
                }
            }
            map
        };

        for (_, state) in combined.into_iter() {
            if let Some(foi) = &state.foi {
                if foi.data.iter().any(|v| v.is_nan()) {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "NaN encountered in foi_data",
                    )
                    .into());
                }

                let foi_mean = mean_kahan(&foi.data);

                let foi_99th_percentile = percentile_nearest(&foi.data, 0.99)?;
                let foi_95th_percentile = percentile_nearest(&foi.data, 0.95)?;
                let foi_90th_percentile = percentile_nearest(&foi.data, 0.90)?;
                let foi_85th_percentile = percentile_nearest(&foi.data, 0.85)?;
                let foi_80th_percentile = percentile_nearest(&foi.data, 0.80)?;
                let foi_75th_percentile = percentile_nearest(&foi.data, 0.75)?;
                let foi_70th_percentile = percentile_nearest(&foi.data, 0.70)?;
                let foi_65th_percentile = percentile_nearest(&foi.data, 0.65)?;
                let foi_60th_percentile = percentile_nearest(&foi.data, 0.60)?;
                let foi_55th_percentile = percentile_nearest(&foi.data, 0.55)?;
                let foi_50th_percentile = percentile_nearest(&foi.data, 0.50)?;
                let foi_45th_percentile = percentile_nearest(&foi.data, 0.45)?;
                let foi_40th_percentile = percentile_nearest(&foi.data, 0.40)?;
                let foi_35th_percentile = percentile_nearest(&foi.data, 0.35)?;
                let foi_30th_percentile = percentile_nearest(&foi.data, 0.30)?;
                let foi_25th_percentile = percentile_nearest(&foi.data, 0.25)?;
                let foi_20th_percentile = percentile_nearest(&foi.data, 0.20)?;
                let foi_15th_percentile = percentile_nearest(&foi.data, 0.15)?;
                let foi_10th_percentile = percentile_nearest(&foi.data, 0.10)?;

                let (
                    foi_entries_above_99th_percentile,
                    foi_entries_above_95th_percentile,
                    foi_entries_above_90th_percentile,
                    foi_entries_above_85th_percentile,
                    foi_entries_above_80th_percentile,
                    foi_entries_above_75th_percentile,
                    foi_entries_above_70th_percentile,
                    foi_entries_above_65th_percentile,
                    foi_entries_above_60th_percentile,
                    foi_entries_above_55th_percentile,
                    foi_entries_above_50th_percentile,
                    foi_entries_above_45th_percentile,
                    foi_entries_above_40th_percentile,
                    foi_entries_above_35th_percentile,
                    foi_entries_above_30th_percentile,
                    foi_entries_above_25th_percentile,
                    foi_entries_above_20th_percentile,
                    foi_entries_above_15th_percentile,
                    foi_entries_above_10th_percentile,
                ) = {
                    let mut foi_entries_above_99th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_95th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_90th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_85th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_80th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_75th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_70th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_65th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_60th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_55th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_50th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_45th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_40th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_35th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_30th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_25th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_20th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_15th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    let mut foi_entries_above_10th_percentile: (Vec<f64>, Vec<f64>) =
                        Default::default();
                    for (index, foi_value) in foi.data.iter().enumerate() {
                        let coordinates = index_to_coordinates(index, width as usize);
                        let coordinates_1_indexed = (coordinates.0 + 1, coordinates.1 + 1);
                        let distance =
                            euclidean_distance(&infection_source, &coordinates_1_indexed) * cd;
                        if distance == 0.0 {
                            continue;
                        }
                        if foi_value >= &foi_99th_percentile {
                            foi_entries_above_99th_percentile.0.push(*foi_value);
                            foi_entries_above_99th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_95th_percentile {
                            foi_entries_above_95th_percentile.0.push(*foi_value);
                            foi_entries_above_95th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_90th_percentile {
                            foi_entries_above_90th_percentile.0.push(*foi_value);
                            foi_entries_above_90th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_85th_percentile {
                            foi_entries_above_85th_percentile.0.push(*foi_value);
                            foi_entries_above_85th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_80th_percentile {
                            foi_entries_above_80th_percentile.0.push(*foi_value);
                            foi_entries_above_80th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_75th_percentile {
                            foi_entries_above_75th_percentile.0.push(*foi_value);
                            foi_entries_above_75th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_70th_percentile {
                            foi_entries_above_70th_percentile.0.push(*foi_value);
                            foi_entries_above_70th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_65th_percentile {
                            foi_entries_above_65th_percentile.0.push(*foi_value);
                            foi_entries_above_65th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_60th_percentile {
                            foi_entries_above_60th_percentile.0.push(*foi_value);
                            foi_entries_above_60th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_55th_percentile {
                            foi_entries_above_55th_percentile.0.push(*foi_value);
                            foi_entries_above_55th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_50th_percentile {
                            foi_entries_above_50th_percentile.0.push(*foi_value);
                            foi_entries_above_50th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_45th_percentile {
                            foi_entries_above_45th_percentile.0.push(*foi_value);
                            foi_entries_above_45th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_40th_percentile {
                            foi_entries_above_40th_percentile.0.push(*foi_value);
                            foi_entries_above_40th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_35th_percentile {
                            foi_entries_above_35th_percentile.0.push(*foi_value);
                            foi_entries_above_35th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_30th_percentile {
                            foi_entries_above_30th_percentile.0.push(*foi_value);
                            foi_entries_above_30th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_25th_percentile {
                            foi_entries_above_25th_percentile.0.push(*foi_value);
                            foi_entries_above_25th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_20th_percentile {
                            foi_entries_above_20th_percentile.0.push(*foi_value);
                            foi_entries_above_20th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_15th_percentile {
                            foi_entries_above_15th_percentile.0.push(*foi_value);
                            foi_entries_above_15th_percentile.1.push(distance);
                        }
                        if foi_value >= &foi_10th_percentile {
                            foi_entries_above_10th_percentile.0.push(*foi_value);
                            foi_entries_above_10th_percentile.1.push(distance);
                        }
                    }
                    (
                        foi_entries_above_99th_percentile,
                        foi_entries_above_95th_percentile,
                        foi_entries_above_90th_percentile,
                        foi_entries_above_85th_percentile,
                        foi_entries_above_80th_percentile,
                        foi_entries_above_75th_percentile,
                        foi_entries_above_70th_percentile,
                        foi_entries_above_65th_percentile,
                        foi_entries_above_60th_percentile,
                        foi_entries_above_55th_percentile,
                        foi_entries_above_50th_percentile,
                        foi_entries_above_45th_percentile,
                        foi_entries_above_40th_percentile,
                        foi_entries_above_35th_percentile,
                        foi_entries_above_30th_percentile,
                        foi_entries_above_25th_percentile,
                        foi_entries_above_20th_percentile,
                        foi_entries_above_15th_percentile,
                        foi_entries_above_10th_percentile,
                    )
                };

                let foi_99th_percentile_mean =
                    mean_kahan(foi_entries_above_99th_percentile.0.iter());
                let foi_95th_percentile_mean =
                    mean_kahan(foi_entries_above_95th_percentile.0.iter());
                let foi_90th_percentile_mean =
                    mean_kahan(foi_entries_above_90th_percentile.0.iter());
                let foi_85th_percentile_mean =
                    mean_kahan(foi_entries_above_85th_percentile.0.iter());
                let foi_80th_percentile_mean =
                    mean_kahan(foi_entries_above_80th_percentile.0.iter());
                let foi_75th_percentile_mean =
                    mean_kahan(foi_entries_above_75th_percentile.0.iter());
                let foi_70th_percentile_mean =
                    mean_kahan(foi_entries_above_70th_percentile.0.iter());
                let foi_65th_percentile_mean =
                    mean_kahan(foi_entries_above_65th_percentile.0.iter());
                let foi_60th_percentile_mean =
                    mean_kahan(foi_entries_above_60th_percentile.0.iter());
                let foi_55th_percentile_mean =
                    mean_kahan(foi_entries_above_55th_percentile.0.iter());
                let foi_50th_percentile_mean =
                    mean_kahan(foi_entries_above_50th_percentile.0.iter());
                let foi_45th_percentile_mean =
                    mean_kahan(foi_entries_above_45th_percentile.0.iter());
                let foi_40th_percentile_mean =
                    mean_kahan(foi_entries_above_40th_percentile.0.iter());
                let foi_35th_percentile_mean =
                    mean_kahan(foi_entries_above_35th_percentile.0.iter());
                let foi_30th_percentile_mean =
                    mean_kahan(foi_entries_above_30th_percentile.0.iter());
                let foi_25th_percentile_mean =
                    mean_kahan(foi_entries_above_25th_percentile.0.iter());
                let foi_20th_percentile_mean =
                    mean_kahan(foi_entries_above_20th_percentile.0.iter());
                let foi_15th_percentile_mean =
                    mean_kahan(foi_entries_above_15th_percentile.0.iter());
                let foi_10th_percentile_mean =
                    mean_kahan(foi_entries_above_10th_percentile.0.iter());

                let foi_99th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_99th_percentile.1.iter());
                let foi_95th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_95th_percentile.1.iter());
                let foi_90th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_90th_percentile.1.iter());
                let foi_85th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_85th_percentile.1.iter());
                let foi_80th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_80th_percentile.1.iter());
                let foi_75th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_75th_percentile.1.iter());
                let foi_70th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_70th_percentile.1.iter());
                let foi_65th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_65th_percentile.1.iter());
                let foi_60th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_60th_percentile.1.iter());
                let foi_55th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_55th_percentile.1.iter());
                let foi_50th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_50th_percentile.1.iter());
                let foi_45th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_45th_percentile.1.iter());
                let foi_40th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_40th_percentile.1.iter());
                let foi_35th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_35th_percentile.1.iter());
                let foi_30th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_30th_percentile.1.iter());
                let foi_25th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_25th_percentile.1.iter());
                let foi_20th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_20th_percentile.1.iter());
                let foi_15th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_15th_percentile.1.iter());
                let foi_10th_percentile_mean_distance_from_source =
                    mean_kahan(foi_entries_above_10th_percentile.1.iter());

                let foi_99th_percentile_spread_rate_mpy =
                    foi_99th_percentile_mean_distance_from_source
                        - previous_foi_99th_percentile_mean_distance_from_source;
                let foi_95th_percentile_spread_rate_mpy =
                    foi_95th_percentile_mean_distance_from_source
                        - previous_foi_95th_percentile_mean_distance_from_source;
                let foi_90th_percentile_spread_rate_mpy =
                    foi_90th_percentile_mean_distance_from_source
                        - previous_foi_90th_percentile_mean_distance_from_source;
                let foi_85th_percentile_spread_rate_mpy =
                    foi_85th_percentile_mean_distance_from_source
                        - previous_foi_85th_percentile_mean_distance_from_source;
                let foi_80th_percentile_spread_rate_mpy =
                    foi_80th_percentile_mean_distance_from_source
                        - previous_foi_80th_percentile_mean_distance_from_source;
                let foi_75th_percentile_spread_rate_mpy =
                    foi_75th_percentile_mean_distance_from_source
                        - previous_foi_75th_percentile_mean_distance_from_source;
                let foi_70th_percentile_spread_rate_mpy =
                    foi_70th_percentile_mean_distance_from_source
                        - previous_foi_70th_percentile_mean_distance_from_source;
                let foi_65th_percentile_spread_rate_mpy =
                    foi_65th_percentile_mean_distance_from_source
                        - previous_foi_65th_percentile_mean_distance_from_source;
                let foi_60th_percentile_spread_rate_mpy =
                    foi_60th_percentile_mean_distance_from_source
                        - previous_foi_60th_percentile_mean_distance_from_source;
                let foi_55th_percentile_spread_rate_mpy =
                    foi_55th_percentile_mean_distance_from_source
                        - previous_foi_55th_percentile_mean_distance_from_source;
                let foi_50th_percentile_spread_rate_mpy =
                    foi_50th_percentile_mean_distance_from_source
                        - previous_foi_50th_percentile_mean_distance_from_source;
                let foi_45th_percentile_spread_rate_mpy =
                    foi_45th_percentile_mean_distance_from_source
                        - previous_foi_45th_percentile_mean_distance_from_source;
                let foi_40th_percentile_spread_rate_mpy =
                    foi_40th_percentile_mean_distance_from_source
                        - previous_foi_40th_percentile_mean_distance_from_source;
                let foi_35th_percentile_spread_rate_mpy =
                    foi_35th_percentile_mean_distance_from_source
                        - previous_foi_35th_percentile_mean_distance_from_source;
                let foi_30th_percentile_spread_rate_mpy =
                    foi_30th_percentile_mean_distance_from_source
                        - previous_foi_30th_percentile_mean_distance_from_source;
                let foi_25th_percentile_spread_rate_mpy =
                    foi_25th_percentile_mean_distance_from_source
                        - previous_foi_25th_percentile_mean_distance_from_source;
                let foi_20th_percentile_spread_rate_mpy =
                    foi_20th_percentile_mean_distance_from_source
                        - previous_foi_20th_percentile_mean_distance_from_source;
                let foi_15th_percentile_spread_rate_mpy =
                    foi_15th_percentile_mean_distance_from_source
                        - previous_foi_15th_percentile_mean_distance_from_source;
                let foi_10th_percentile_spread_rate_mpy =
                    foi_10th_percentile_mean_distance_from_source
                        - previous_foi_10th_percentile_mean_distance_from_source;

                previous_foi_99th_percentile_mean_distance_from_source =
                    foi_99th_percentile_mean_distance_from_source;
                previous_foi_95th_percentile_mean_distance_from_source =
                    foi_95th_percentile_mean_distance_from_source;
                previous_foi_90th_percentile_mean_distance_from_source =
                    foi_90th_percentile_mean_distance_from_source;
                previous_foi_85th_percentile_mean_distance_from_source =
                    foi_85th_percentile_mean_distance_from_source;
                previous_foi_80th_percentile_mean_distance_from_source =
                    foi_80th_percentile_mean_distance_from_source;
                previous_foi_75th_percentile_mean_distance_from_source =
                    foi_75th_percentile_mean_distance_from_source;
                previous_foi_70th_percentile_mean_distance_from_source =
                    foi_70th_percentile_mean_distance_from_source;
                previous_foi_65th_percentile_mean_distance_from_source =
                    foi_65th_percentile_mean_distance_from_source;
                previous_foi_60th_percentile_mean_distance_from_source =
                    foi_60th_percentile_mean_distance_from_source;
                previous_foi_55th_percentile_mean_distance_from_source =
                    foi_55th_percentile_mean_distance_from_source;
                previous_foi_50th_percentile_mean_distance_from_source =
                    foi_50th_percentile_mean_distance_from_source;
                previous_foi_45th_percentile_mean_distance_from_source =
                    foi_45th_percentile_mean_distance_from_source;
                previous_foi_40th_percentile_mean_distance_from_source =
                    foi_40th_percentile_mean_distance_from_source;
                previous_foi_35th_percentile_mean_distance_from_source =
                    foi_35th_percentile_mean_distance_from_source;
                previous_foi_30th_percentile_mean_distance_from_source =
                    foi_30th_percentile_mean_distance_from_source;
                previous_foi_25th_percentile_mean_distance_from_source =
                    foi_25th_percentile_mean_distance_from_source;
                previous_foi_20th_percentile_mean_distance_from_source =
                    foi_20th_percentile_mean_distance_from_source;
                previous_foi_15th_percentile_mean_distance_from_source =
                    foi_15th_percentile_mean_distance_from_source;
                previous_foi_10th_percentile_mean_distance_from_source =
                    foi_10th_percentile_mean_distance_from_source;

                workbook.write_number(Column::FOIMean.number(), foi_mean)?;
                workbook.write_number(Column::FOI99thPercentile.number(), foi_99th_percentile)?;
                workbook.write_number(Column::FOI95thPercentile.number(), foi_95th_percentile)?;
                workbook.write_number(Column::FOI90thPercentile.number(), foi_90th_percentile)?;
                workbook.write_number(Column::FOI85thPercentile.number(), foi_85th_percentile)?;
                workbook.write_number(Column::FOI80thPercentile.number(), foi_80th_percentile)?;
                workbook.write_number(Column::FOI75thPercentile.number(), foi_75th_percentile)?;
                workbook.write_number(Column::FOI70thPercentile.number(), foi_70th_percentile)?;
                workbook.write_number(Column::FOI65thPercentile.number(), foi_65th_percentile)?;
                workbook.write_number(Column::FOI60thPercentile.number(), foi_60th_percentile)?;
                workbook.write_number(Column::FOI55thPercentile.number(), foi_55th_percentile)?;
                workbook.write_number(Column::FOI50thPercentile.number(), foi_50th_percentile)?;
                workbook.write_number(Column::FOI45thPercentile.number(), foi_45th_percentile)?;
                workbook.write_number(Column::FOI40thPercentile.number(), foi_40th_percentile)?;
                workbook.write_number(Column::FOI35thPercentile.number(), foi_35th_percentile)?;
                workbook.write_number(Column::FOI30thPercentile.number(), foi_30th_percentile)?;
                workbook.write_number(Column::FOI25thPercentile.number(), foi_25th_percentile)?;
                workbook.write_number(Column::FOI20thPercentile.number(), foi_20th_percentile)?;
                workbook.write_number(Column::FOI15thPercentile.number(), foi_15th_percentile)?;
                workbook.write_number(Column::FOI10thPercentile.number(), foi_10th_percentile)?;
                workbook.write_number(
                    Column::FOI99thPercentileMean.number(),
                    foi_99th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI95thPercentileMean.number(),
                    foi_95th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI90thPercentileMean.number(),
                    foi_90th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI85thPercentileMean.number(),
                    foi_85th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI80thPercentileMean.number(),
                    foi_80th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI75thPercentileMean.number(),
                    foi_75th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI70thPercentileMean.number(),
                    foi_70th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI65thPercentileMean.number(),
                    foi_65th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI60thPercentileMean.number(),
                    foi_60th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI55thPercentileMean.number(),
                    foi_55th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI50thPercentileMean.number(),
                    foi_50th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI45thPercentileMean.number(),
                    foi_45th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI40thPercentileMean.number(),
                    foi_40th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI35thPercentileMean.number(),
                    foi_35th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI30thPercentileMean.number(),
                    foi_30th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI25thPercentileMean.number(),
                    foi_25th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI20thPercentileMean.number(),
                    foi_20th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI15thPercentileMean.number(),
                    foi_15th_percentile_mean,
                )?;
                workbook.write_number(
                    Column::FOI10thPercentileMean.number(),
                    foi_10th_percentile_mean,
                )?;

                workbook.write_number(
                    Column::FOI99thPercentileMeanDistanceMeters.number(),
                    foi_99th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI95thPercentileMeanDistanceMeters.number(),
                    foi_95th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI90thPercentileMeanDistanceMeters.number(),
                    foi_90th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI85thPercentileMeanDistanceMeters.number(),
                    foi_85th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI80thPercentileMeanDistanceMeters.number(),
                    foi_80th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI75thPercentileMeanDistanceMeters.number(),
                    foi_75th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI70thPercentileMeanDistanceMeters.number(),
                    foi_70th_percentile_mean_distance_from_source,
                )?;

                workbook.write_number(
                    Column::FOI65thPercentileMeanDistanceMeters.number(),
                    foi_65th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI60thPercentileMeanDistanceMeters.number(),
                    foi_60th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI55thPercentileMeanDistanceMeters.number(),
                    foi_55th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI50thPercentileMeanDistanceMeters.number(),
                    foi_50th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI45thPercentileMeanDistanceMeters.number(),
                    foi_45th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI40thPercentileMeanDistanceMeters.number(),
                    foi_40th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI35thPercentileMeanDistanceMeters.number(),
                    foi_35th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI30thPercentileMeanDistanceMeters.number(),
                    foi_30th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI25thPercentileMeanDistanceMeters.number(),
                    foi_25th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI20thPercentileMeanDistanceMeters.number(),
                    foi_20th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI15thPercentileMeanDistanceMeters.number(),
                    foi_15th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI10thPercentileMeanDistanceMeters.number(),
                    foi_10th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::FOI99thPercentileSpreadMetersPerYear.number(),
                    foi_99th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI95thPercentileSpreadMetersPerYear.number(),
                    foi_95th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI90thPercentileSpreadMetersPerYear.number(),
                    foi_90th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI85thPercentileSpreadMetersPerYear.number(),
                    foi_85th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI80thPercentileSpreadMetersPerYear.number(),
                    foi_80th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI75thPercentileSpreadMetersPerYear.number(),
                    foi_75th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI70thPercentileSpreadMetersPerYear.number(),
                    foi_70th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI65thPercentileSpreadMetersPerYear.number(),
                    foi_65th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI60thPercentileSpreadMetersPerYear.number(),
                    foi_60th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI55thPercentileSpreadMetersPerYear.number(),
                    foi_55th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI50thPercentileSpreadMetersPerYear.number(),
                    foi_50th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI45thPercentileSpreadMetersPerYear.number(),
                    foi_45th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI40thPercentileSpreadMetersPerYear.number(),
                    foi_40th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI35thPercentileSpreadMetersPerYear.number(),
                    foi_35th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI30thPercentileSpreadMetersPerYear.number(),
                    foi_30th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI25thPercentileSpreadMetersPerYear.number(),
                    foi_25th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI20thPercentileSpreadMetersPerYear.number(),
                    foi_20th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI15thPercentileSpreadMetersPerYear.number(),
                    foi_15th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::FOI10thPercentileSpreadMetersPerYear.number(),
                    foi_10th_percentile_spread_rate_mpy,
                )?;
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
                    width,
                    height,
                    &normalized,
                    &cfg,
                )?;
            }

            if let Some(infection) = &state.infection {
                let infection_distances =
                    distances_from_infection_source_1_indexed
                        .iter()
                        .filter(|&(coordinates, _)| {
                            infection
                                .infected_sites
                                .contains(&(coordinates.0 as u32, coordinates.1 as u32))
                        });
                /* let infection_distances = {
                    let mut map: HashMap<(usize, usize), f64> = HashMap::new();
                    for &(x, y) in infection.infected_sites.iter() {
                        let site = (x as usize, y as usize);
                        if let Some(distance) = distances_from_infection_source_1_indexed.get(&site)
                        {
                            map.insert(site, *distance);
                        }
                    }
                    map
                }; */

                let total_number_of_infected_sites = infection.infected_sites.len();
                let newly_infected_sites =
                    infection.infected_sites.len() - previous_number_of_infected_sites;
                let infected_area = infection.infected_sites.len() as f64
                    / (width as usize * height as usize) as f64;

                let values = infection_distances
                    .clone()
                    .map(|(_, distance)| *distance)
                    .collect::<Vec<f64>>();
                let (
                    infection_99th_percentile_distance,
                    infection_95th_percentile_distance,
                    infection_90th_percentile_distance,
                    infection_85th_percentile_distance,
                    infection_80th_percentile_distance,
                    infection_75th_percentile_distance,
                    infection_70th_percentile_distance,
                    infection_65th_percentile_distance,
                    infection_60th_percentile_distance,
                    infection_55th_percentile_distance,
                    infection_50th_percentile_distance,
                    infection_45th_percentile_distance,
                    infection_40th_percentile_distance,
                    infection_35th_percentile_distance,
                    infection_30th_percentile_distance,
                    infection_25th_percentile_distance,
                    infection_20th_percentile_distance,
                    infection_15th_percentile_distance,
                    infection_10th_percentile_distance,
                ) = if values.len() > 0 {
                    (
                        percentile_nearest(&values, 0.99)?,
                        percentile_nearest(&values, 0.95)?,
                        percentile_nearest(&values, 0.90)?,
                        percentile_nearest(&values, 0.85)?,
                        percentile_nearest(&values, 0.80)?,
                        percentile_nearest(&values, 0.75)?,
                        percentile_nearest(&values, 0.70)?,
                        percentile_nearest(&values, 0.65)?,
                        percentile_nearest(&values, 0.60)?,
                        percentile_nearest(&values, 0.55)?,
                        percentile_nearest(&values, 0.50)?,
                        percentile_nearest(&values, 0.45)?,
                        percentile_nearest(&values, 0.40)?,
                        percentile_nearest(&values, 0.35)?,
                        percentile_nearest(&values, 0.30)?,
                        percentile_nearest(&values, 0.25)?,
                        percentile_nearest(&values, 0.20)?,
                        percentile_nearest(&values, 0.15)?,
                        percentile_nearest(&values, 0.10)?,
                    )
                } else {
                    (
                        0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
                        0.0, 0.0, 0.0, 0.0,
                    )
                };

                let infection_mean_distance_from_source = mean_kahan(values.iter());

                let infection_99th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_99th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };

                let infection_95th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_95th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };

                let infection_90th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_90th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_85th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_85th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };

                let infection_80th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_80th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };

                let infection_75th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_75th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };

                let infection_70th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_70th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_65th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_65th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_60th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_60th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_55th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_55th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_50th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_50th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_45th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_45th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_40th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_40th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_35th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_35th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_30th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_30th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_25th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_25th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_20th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_20th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_15th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_15th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };
                let infection_10th_percentile_mean_distance_from_source = {
                    let iter = infection_distances
                        .clone()
                        .filter(|(_, distance)| *distance >= &infection_10th_percentile_distance)
                        .map(|t| t.1);
                    mean_kahan(iter)
                };

                let infection_99th_percentile_spread_rate_mpy =
                    infection_99th_percentile_mean_distance_from_source
                        - previous_infection_99th_percentile_mean_distance_from_source;
                let infection_95th_percentile_spread_rate_mpy =
                    infection_95th_percentile_mean_distance_from_source
                        - previous_infection_95th_percentile_mean_distance_from_source;
                let infection_90th_percentile_spread_rate_mpy =
                    infection_90th_percentile_mean_distance_from_source
                        - previous_infection_90th_percentile_mean_distance_from_source;
                let infection_85th_percentile_spread_rate_mpy =
                    infection_85th_percentile_mean_distance_from_source
                        - previous_infection_85th_percentile_mean_distance_from_source;
                let infection_80th_percentile_spread_rate_mpy =
                    infection_80th_percentile_mean_distance_from_source
                        - previous_infection_80th_percentile_mean_distance_from_source;
                let infection_75th_percentile_spread_rate_mpy =
                    infection_75th_percentile_mean_distance_from_source
                        - previous_infection_75th_percentile_mean_distance_from_source;
                let infection_70th_percentile_spread_rate_mpy =
                    infection_70th_percentile_mean_distance_from_source
                        - previous_infection_70th_percentile_mean_distance_from_source;
                let infection_65th_percentile_spread_rate_mpy =
                    infection_65th_percentile_mean_distance_from_source
                        - previous_infection_65th_percentile_mean_distance_from_source;
                let infection_60th_percentile_spread_rate_mpy =
                    infection_60th_percentile_mean_distance_from_source
                        - previous_infection_60th_percentile_mean_distance_from_source;
                let infection_55th_percentile_spread_rate_mpy =
                    infection_55th_percentile_mean_distance_from_source
                        - previous_infection_55th_percentile_mean_distance_from_source;
                let infection_50th_percentile_spread_rate_mpy =
                    infection_50th_percentile_mean_distance_from_source
                        - previous_infection_50th_percentile_mean_distance_from_source;
                let infection_45th_percentile_spread_rate_mpy =
                    infection_45th_percentile_mean_distance_from_source
                        - previous_infection_45th_percentile_mean_distance_from_source;
                let infection_40th_percentile_spread_rate_mpy =
                    infection_40th_percentile_mean_distance_from_source
                        - previous_infection_40th_percentile_mean_distance_from_source;
                let infection_35th_percentile_spread_rate_mpy =
                    infection_35th_percentile_mean_distance_from_source
                        - previous_infection_35th_percentile_mean_distance_from_source;
                let infection_30th_percentile_spread_rate_mpy =
                    infection_30th_percentile_mean_distance_from_source
                        - previous_infection_30th_percentile_mean_distance_from_source;
                let infection_25th_percentile_spread_rate_mpy =
                    infection_25th_percentile_mean_distance_from_source
                        - previous_infection_25th_percentile_mean_distance_from_source;
                let infection_20th_percentile_spread_rate_mpy =
                    infection_20th_percentile_mean_distance_from_source
                        - previous_infection_20th_percentile_mean_distance_from_source;
                let infection_15th_percentile_spread_rate_mpy =
                    infection_15th_percentile_mean_distance_from_source
                        - previous_infection_15th_percentile_mean_distance_from_source;
                let infection_10th_percentile_spread_rate_mpy =
                    infection_10th_percentile_mean_distance_from_source
                        - previous_infection_10th_percentile_mean_distance_from_source;

                previous_infection_99th_percentile_mean_distance_from_source =
                    infection_99th_percentile_mean_distance_from_source;
                previous_infection_95th_percentile_mean_distance_from_source =
                    infection_95th_percentile_mean_distance_from_source;
                previous_infection_90th_percentile_mean_distance_from_source =
                    infection_90th_percentile_mean_distance_from_source;
                previous_infection_85th_percentile_mean_distance_from_source =
                    infection_85th_percentile_mean_distance_from_source;
                previous_infection_80th_percentile_mean_distance_from_source =
                    infection_80th_percentile_mean_distance_from_source;
                previous_infection_75th_percentile_mean_distance_from_source =
                    infection_75th_percentile_mean_distance_from_source;
                previous_infection_70th_percentile_mean_distance_from_source =
                    infection_70th_percentile_mean_distance_from_source;
                previous_infection_65th_percentile_mean_distance_from_source =
                    infection_65th_percentile_mean_distance_from_source;
                previous_infection_60th_percentile_mean_distance_from_source =
                    infection_60th_percentile_mean_distance_from_source;
                previous_infection_55th_percentile_mean_distance_from_source =
                    infection_55th_percentile_mean_distance_from_source;
                previous_infection_50th_percentile_mean_distance_from_source =
                    infection_50th_percentile_mean_distance_from_source;
                previous_infection_45th_percentile_mean_distance_from_source =
                    infection_45th_percentile_mean_distance_from_source;
                previous_infection_40th_percentile_mean_distance_from_source =
                    infection_40th_percentile_mean_distance_from_source;
                previous_infection_35th_percentile_mean_distance_from_source =
                    infection_35th_percentile_mean_distance_from_source;
                previous_infection_30th_percentile_mean_distance_from_source =
                    infection_30th_percentile_mean_distance_from_source;
                previous_infection_25th_percentile_mean_distance_from_source =
                    infection_25th_percentile_mean_distance_from_source;
                previous_infection_20th_percentile_mean_distance_from_source =
                    infection_20th_percentile_mean_distance_from_source;
                previous_infection_15th_percentile_mean_distance_from_source =
                    infection_15th_percentile_mean_distance_from_source;
                previous_infection_10th_percentile_mean_distance_from_source =
                    infection_10th_percentile_mean_distance_from_source;

                let (
                    newly_infected_sites_change_modifier,
                    newly_infected_sites_change,
                    newly_infected_sites_change_percent,
                ) = if previous_newly_infected_sites > 0 {
                    (
                        newly_infected_sites as f64 / previous_newly_infected_sites as f64,
                        newly_infected_sites as f64 - previous_newly_infected_sites as f64,
                        ((newly_infected_sites as f64 - previous_newly_infected_sites as f64)
                            / previous_newly_infected_sites as f64)
                            * 100.0,
                    )
                } else {
                    (1.0, 0.0, 0.0)
                };

                let (
                    infection_area_change_modifier,
                    infection_area_change,
                    infection_area_change_percent,
                ) = if previous_infected_area > 0.0 {
                    (
                        infected_area / previous_infected_area,
                        infected_area - previous_infected_area,
                        ((infected_area as f64 - previous_infected_area as f64)
                            / previous_infected_area as f64)
                            * 100.0,
                    )
                } else {
                    (1.0, 0.0, 0.0)
                };

                if let Some(biomass) = &state.biomass {
                    let (total_healthy_biomass, total_infected_biomass, total_ignored_biomass) =
                        biomass.data.iter().fold(
                            (0u64, 0u64, 0u64),
                            |(sum_healthy, sum_infected, sum_ignored),
                             &(healthy, infected, ignored)| {
                                (
                                    sum_healthy.saturating_add(healthy),
                                    sum_infected.saturating_add(infected),
                                    sum_ignored.saturating_add(ignored),
                                )
                            },
                        );
                    let (total_healthy_biomass, total_infected_biomass, total_ignored_biomass) = (
                        total_healthy_biomass as f64,
                        total_infected_biomass as f64,
                        total_ignored_biomass as f64,
                    );

                    let total_biomass =
                        total_healthy_biomass + total_infected_biomass + total_ignored_biomass;

                    let proportion_infected_biomass = total_infected_biomass / total_biomass;
                    let proportion_host_infected_biomass =
                        total_infected_biomass / (total_healthy_biomass + total_infected_biomass);

                    let (
                        total_healthy_biomass_change_modifier,
                        total_healthy_biomass_change,
                        total_healthy_biomass_change_percent,
                    ) = if previous_total_healthy_biomass > 0.0 {
                        (
                            total_healthy_biomass / previous_total_healthy_biomass,
                            total_healthy_biomass - previous_total_healthy_biomass,
                            ((total_healthy_biomass as f64
                                - previous_total_healthy_biomass as f64)
                                / previous_total_healthy_biomass as f64)
                                * 100.0,
                        )
                    } else {
                        (1.0, 0.0, 0.0)
                    };
                    let (
                        total_infected_biomass_change_modifier,
                        total_infected_biomass_change,
                        total_infected_biomass_change_percent,
                    ) = if previous_total_infected_biomass > 0.0 {
                        (
                            total_infected_biomass / previous_total_infected_biomass,
                            total_infected_biomass - previous_total_infected_biomass,
                            ((total_infected_biomass as f64
                                - previous_total_infected_biomass as f64)
                                / previous_total_infected_biomass as f64)
                                * 100.0,
                        )
                    } else {
                        (1.0, 0.0, 0.0)
                    };
                    let (
                        total_ignored_biomass_change_modifier,
                        total_ignored_biomass_change,
                        total_ignored_biomass_change_percent,
                    ) = if previous_total_ignored_biomass > 0.0 {
                        (
                            total_ignored_biomass / previous_total_ignored_biomass,
                            total_ignored_biomass - previous_total_ignored_biomass,
                            ((total_ignored_biomass as f64
                                - previous_total_ignored_biomass as f64)
                                / previous_total_ignored_biomass as f64)
                                * 100.0,
                        )
                    } else {
                        (1.0, 0.0, 0.0)
                    };
                    let (
                        total_biomass_change_modifier,
                        total_biomass_change,
                        total_biomass_change_percent,
                    ) = if previous_total_biomass > 0.0 {
                        (
                            total_biomass / previous_total_biomass,
                            total_biomass - previous_total_biomass,
                            ((total_biomass as f64 - previous_total_biomass as f64)
                                / previous_total_biomass as f64)
                                * 100.0,
                        )
                    } else {
                        (1.0, 0.0, 0.0)
                    };

                    previous_total_healthy_biomass = total_healthy_biomass;
                    previous_total_infected_biomass = total_infected_biomass;
                    previous_total_ignored_biomass = total_ignored_biomass;
                    previous_total_biomass = total_biomass;

                    workbook.write_number(
                        Column::TotalHealthyBiomass.number(),
                        total_healthy_biomass,
                    )?;
                    workbook.write_number(
                        Column::TotalInfectedBiomass.number(),
                        total_infected_biomass,
                    )?;
                    workbook.write_number(
                        Column::TotalIgnoredBiomass.number(),
                        total_ignored_biomass,
                    )?;
                    workbook.write_number(Column::TotalBiomass.number(), total_biomass)?;
                    workbook.write_number(
                        Column::ProportionInfectedBiomass.number(),
                        proportion_infected_biomass,
                    )?;
                    workbook.write_number(
                        Column::ProportionHostInfectedBiomass.number(),
                        proportion_host_infected_biomass,
                    )?;
                    workbook.write_number(
                        Column::TotalHealthyBiomassChange.number(),
                        total_healthy_biomass_change,
                    )?;
                    workbook.write_number(
                        Column::TotalInfectedBiomassChange.number(),
                        total_infected_biomass_change,
                    )?;
                    workbook.write_number(
                        Column::TotalIgnoredBiomassChange.number(),
                        total_ignored_biomass_change,
                    )?;
                    workbook
                        .write_number(Column::TotalBiomassChange.number(), total_biomass_change)?;
                    workbook.write_number(
                        Column::TotalHealthyBiomassChangePercentage.number(),
                        total_healthy_biomass_change_percent,
                    )?;
                    workbook.write_number(
                        Column::TotalInfectedBiomassChangePercentage.number(),
                        total_infected_biomass_change_percent,
                    )?;
                    workbook.write_number(
                        Column::TotalIgnoredBiomassChangePercentage.number(),
                        total_ignored_biomass_change_percent,
                    )?;
                    workbook.write_number(
                        Column::TotalBiomassChangePercentage.number(),
                        total_biomass_change_percent,
                    )?;
                    workbook.write_number(
                        Column::TotalHealthyBiomassChangeMod.number(),
                        total_healthy_biomass_change_modifier,
                    )?;
                    workbook.write_number(
                        Column::TotalInfectedBiomassChangeMod.number(),
                        total_infected_biomass_change_modifier,
                    )?;
                    workbook.write_number(
                        Column::TotalIgnoredBiomassChangeMod.number(),
                        total_ignored_biomass_change_modifier,
                    )?;
                    workbook.write_number(
                        Column::TotalBiomassChangeMod.number(),
                        total_biomass_change_modifier,
                    )?;
                }

                previous_number_of_infected_sites = infection.infected_sites.len();
                previous_infected_area = infected_area;
                previous_newly_infected_sites = newly_infected_sites;

                workbook.write_number(
                    Column::InfTotal.number(),
                    total_number_of_infected_sites as f64,
                )?;
                workbook.write_number(
                    Column::InfMeanDistMeters.number(),
                    infection_mean_distance_from_source,
                )?;
                workbook.write_number(Column::InfNew.number(), newly_infected_sites as f64)?;
                workbook.write_number(
                    Column::InfectedNewChange.number(),
                    newly_infected_sites_change,
                )?;
                workbook.write_number(
                    Column::InfNewChangePercentage.number(),
                    newly_infected_sites_change_percent,
                )?;
                workbook.write_number(
                    Column::InfNewChangeMod.number(),
                    newly_infected_sites_change_modifier,
                )?;
                workbook.write_number(Column::InfAreaSquareMeters.number(), infected_area)?;
                workbook.write_number(
                    Column::InfAreaChangeSquareMeters.number(),
                    infection_area_change,
                )?;
                workbook.write_number(
                    Column::InfAreaChangeSquareMetersPercentage.number(),
                    infection_area_change_percent,
                )?;
                workbook.write_number(
                    Column::InfectedAreaChangeSquareMetersMod.number(),
                    infection_area_change_modifier,
                )?;
                workbook.write_number(
                    Column::Inf99thPercentile.number(),
                    infection_99th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf95thPercentile.number(),
                    infection_95th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf90thPercentile.number(),
                    infection_90th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf85thPercentile.number(),
                    infection_85th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf80thPercentile.number(),
                    infection_80th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf75thPercentile.number(),
                    infection_75th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf70thPercentile.number(),
                    infection_70th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf65thPercentile.number(),
                    infection_65th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf60thPercentile.number(),
                    infection_60th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf55thPercentile.number(),
                    infection_55th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf50thPercentile.number(),
                    infection_50th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf45thPercentile.number(),
                    infection_45th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf40thPercentile.number(),
                    infection_40th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf35thPercentile.number(),
                    infection_35th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf30thPercentile.number(),
                    infection_30th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf25thPercentile.number(),
                    infection_25th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf20thPercentile.number(),
                    infection_20th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf15thPercentile.number(),
                    infection_15th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf10thPercentile.number(),
                    infection_10th_percentile_distance,
                )?;
                workbook.write_number(
                    Column::Inf99thPercentileMeanDistanceMeters.number(),
                    infection_99th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf95thPercentileMeanDistanceMeters.number(),
                    infection_95th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf90thPercentileMeanDistanceMeters.number(),
                    infection_90th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf85thPercentileMeanDistanceMeters.number(),
                    infection_85th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf80thPercentileMeanDistanceMeters.number(),
                    infection_80th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf75thPercentileMeanDistanceMeters.number(),
                    infection_75th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf70thPercentileMeanDistanceMeters.number(),
                    infection_70th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf65thPercentileMeanDistanceMeters.number(),
                    infection_65th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf60thPercentileMeanDistanceMeters.number(),
                    infection_60th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf55thPercentileMeanDistanceMeters.number(),
                    infection_55th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf50thPercentileMeanDistanceMeters.number(),
                    infection_50th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf45thPercentileMeanDistanceMeters.number(),
                    infection_45th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf40thPercentileMeanDistanceMeters.number(),
                    infection_40th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf35thPercentileMeanDistanceMeters.number(),
                    infection_35th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf30thPercentileMeanDistanceMeters.number(),
                    infection_30th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf25thPercentileMeanDistanceMeters.number(),
                    infection_25th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf20thPercentileMeanDistanceMeters.number(),
                    infection_20th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf15thPercentileMeanDistanceMeters.number(),
                    infection_15th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf10thPercentileMeanDistanceMeters.number(),
                    infection_10th_percentile_mean_distance_from_source,
                )?;
                workbook.write_number(
                    Column::Inf99thPercentileSpreadMetersPerYear.number(),
                    infection_99th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf95thPercentileSpreadMetersPerYear.number(),
                    infection_95th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf90thPercentileSpreadMetersPerYear.number(),
                    infection_90th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf85thPercentileSpreadMetersPerYear.number(),
                    infection_85th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf80thPercentileSpreadMetersPerYear.number(),
                    infection_80th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf75thPercentileSpreadMetersPerYear.number(),
                    infection_75th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf70thPercentileSpreadMetersPerYear.number(),
                    infection_70th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf65thPercentileSpreadMetersPerYear.number(),
                    infection_65th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf60thPercentileSpreadMetersPerYear.number(),
                    infection_60th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf55thPercentileSpreadMetersPerYear.number(),
                    infection_55th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf50thPercentileSpreadMetersPerYear.number(),
                    infection_50th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf45thPercentileSpreadMetersPerYear.number(),
                    infection_45th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf40thPercentileSpreadMetersPerYear.number(),
                    infection_40th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf35thPercentileSpreadMetersPerYear.number(),
                    infection_35th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf30thPercentileSpreadMetersPerYear.number(),
                    infection_30th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf25thPercentileSpreadMetersPerYear.number(),
                    infection_25th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf20thPercentileSpreadMetersPerYear.number(),
                    infection_20th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf15thPercentileSpreadMetersPerYear.number(),
                    infection_15th_percentile_spread_rate_mpy,
                )?;
                workbook.write_number(
                    Column::Inf10thPercentileSpreadMetersPerYear.number(),
                    infection_10th_percentile_spread_rate_mpy,
                )?;
            }

            if let Some(mortality) = &state.mortality {
                let total_annual_mortality = mortality.data.iter().sum::<u32>() as f64;
                mortality_values.push(total_annual_mortality);

                workbook.write_number(
                    Column::TotalAnnualMortality.number(),
                    total_annual_mortality,
                )?;
            }

            if let Some(infection) = &state.infection {
                let start = std::time::Instant::now();
                let _ = render_infection_state_png(
                    &args.output_dir,
                    state.timestep,
                    width,
                    height,
                    &infection.healthy_sites,
                    &infection.infected_sites,
                    &infection.ignored_sites,
                    &image_cfg,
                )?;
                let ms = start.elapsed().as_millis();
                img_time_sum_ms += ms;
                img_time_count += 1;
            }

            if let Some(mortality_occurred) = &state.mortality_occurred {
                render_state_map_png(
                    &args.output_dir,
                    "mortality_occurred",
                    state.timestep,
                    width,
                    height,
                    mortality_occurred.data.as_ref(),
                    [255, 0, 0, 255],
                    [0, 0, 0, 255],
                    &image_cfg,
                )?;
            }

            if let Some(infection_occurred) = &state.infection_occurred {
                render_state_map_png(
                    &args.output_dir,
                    "infection_occurred",
                    state.timestep,
                    width,
                    height,
                    infection_occurred.data.as_ref(),
                    [255, 0, 0, 255],
                    [0, 0, 0, 255],
                    &image_cfg,
                )?;
            }

            workbook.write_number(Column::Timestep.number(), state.timestep as f64)?;
            ROW_COUNTER.fetch_add(1, AtomicOrdering::SeqCst);
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

        for column in Column::iter() {
            if let Some(formatter) = column.conditional_formatter() {
                match formatter {
                    ConditionalFormatter::ChangeAbsolute => {
                        workbook.add_conditional_format(column.number(), &negative_condition)?;
                        workbook.add_conditional_format(column.number(), &positive_condition)?;
                        workbook.add_conditional_format(column.number(), &zero_condition)?;
                    }
                    ConditionalFormatter::ChangePercentage => {
                        workbook.add_conditional_format(column.number(), &negative_condition)?;
                        workbook.add_conditional_format(column.number(), &positive_condition)?;
                        workbook.add_conditional_format(column.number(), &zero_condition)?;
                    }
                    ConditionalFormatter::ChangeModifier => {
                        workbook.add_conditional_format(
                            column.number(),
                            &modifier_less_than_one_condition,
                        )?;
                        workbook.add_conditional_format(
                            column.number(),
                            &modifier_greater_than_one_condition,
                        )?;
                        workbook.add_conditional_format(
                            column.number(),
                            &modifier_equal_one_condition,
                        )?;
                    }
                }
            }
        }

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
            workbook.add_conditional_format(
                Column::TotalAnnualMortality.number(),
                &mortality_min_condition,
            )?;
            workbook.add_conditional_format(
                Column::TotalAnnualMortality.number(),
                &mortality_max_condition,
            )?;
            workbook.add_conditional_format(
                Column::TotalAnnualMortality.number(),
                &mortality_median_condition,
            )?;
        }

        let xlsx_path = args.output_dir.join("output.xlsx");
        workbook.save(&xlsx_path.to_string_lossy())?;
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
