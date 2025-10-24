use std::fs::{self, File};
use std::io::{BufWriter, Write};
use std::path::{Path, PathBuf};

pub struct GridRenderConfig {
    pub preferred_cell_px: u32,
    pub max_image_size_px: u32,
    pub background: Option<[u8; 4]>,
}

impl Default for GridRenderConfig {
    fn default() -> Self {
        GridRenderConfig {
            preferred_cell_px: 10,
            max_image_size_px: 16384,
            background: None,
        }
    }
}

pub struct CellCtx {
    pub scale: u32,
    pub cell_w_px: u32,
    pub cell_h_px: u32,
    pub grid_w: u32,
    pub grid_h: u32,
    pub img_w: u32,
    pub img_h: u32,
}

pub fn compute_scale(grid_w: u32, grid_h: u32, cfg: &GridRenderConfig) -> Option<(u32, u32, u32)> {
    if grid_w == 0 || grid_h == 0 {
        return None;
    }
    let mut scale = cfg.preferred_cell_px.max(1);
    while scale > 0
        && (scale.saturating_mul(grid_w) > cfg.max_image_size_px
            || scale.saturating_mul(grid_h) > cfg.max_image_size_px)
    {
        scale -= 1;
    }
    if scale == 0 {
        return None;
    }
    let img_w = grid_w * scale;
    let img_h = grid_h * scale;
    Some((scale, img_w, img_h))
}

fn new_png_encoder<W: Write>(
    w: W,
    width: u32,
    height: u32,
) -> Result<png::Writer<W>, png::EncodingError> {
    let mut encoder = png::Encoder::new(w, width, height);
    encoder.set_color(png::ColorType::Rgba);
    encoder.set_depth(png::BitDepth::Eight);
    encoder.set_compression(png::Compression::default());
    encoder.write_header()
}

fn new_png_encoder_gray16<W: Write>(
    w: W,
    width: u32,
    height: u32,
) -> Result<png::Writer<W>, png::EncodingError> {
    let mut encoder = png::Encoder::new(w, width, height);
    encoder.set_color(png::ColorType::Grayscale);
    encoder.set_depth(png::BitDepth::Sixteen);
    encoder.set_compression(png::Compression::default());
    encoder.write_header()
}

pub fn render_cells_png<F>(
    output_path: impl AsRef<Path>,
    grid_w: u32,
    grid_h: u32,
    cfg: &GridRenderConfig,
    mut painter: F,
) -> Result<(), Box<dyn std::error::Error>>
where
    F: FnMut(u32, u32, &CellCtx) -> [u8; 4],
{
    let (scale, img_w, img_h) = compute_scale(grid_w, grid_h, cfg)
        .ok_or_else(|| format!("grid too large for max image size"))?;
    if let Some(parent) = output_path.as_ref().parent() {
        fs::create_dir_all(parent)?;
    }

    let file = File::create(&output_path)?;
    let mut writer = BufWriter::new(file);
    let mut png_writer = new_png_encoder(&mut writer, img_w, img_h)?;
    let mut stream = png_writer.stream_writer()?;
    let row_bytes = (img_w as usize) * 4;
    let mut row_buf = vec![0u8; row_bytes];

    let ctx = CellCtx {
        scale,
        cell_w_px: scale,
        cell_h_px: scale,
        grid_w,
        grid_h,
        img_w,
        img_h,
    };

    for gy in 0..grid_h {
        for _sy in 0..scale {
            let mut offset = 0usize;
            for gx in 0..grid_w {
                let rgba = painter(gx, gy, &ctx);
                for _ in 0..scale {
                    row_buf[offset] = rgba[0];
                    row_buf[offset + 1] = rgba[1];
                    row_buf[offset + 2] = rgba[2];
                    row_buf[offset + 3] = rgba[3];
                    offset += 4;
                }
            }
            stream.write_all(&row_buf)?;
        }
    }
    stream.finish()?;
    Ok(())
}

pub fn render_pixels_png<F>(
    output_path: impl AsRef<Path>,
    grid_w: u32,
    grid_h: u32,
    cfg: &GridRenderConfig,
    mut painter: F,
) -> Result<(), Box<dyn std::error::Error>>
where
    F: FnMut(u32, u32, u32, u32, &CellCtx) -> [u8; 4],
{
    let (scale, img_w, img_h) = compute_scale(grid_w, grid_h, cfg)
        .ok_or_else(|| format!("grid too large for max image size"))?;
    if let Some(parent) = output_path.as_ref().parent() {
        fs::create_dir_all(parent)?;
    }

    let file = File::create(&output_path)?;
    let mut writer = BufWriter::new(file);
    let mut png_writer = new_png_encoder(&mut writer, img_w, img_h)?;
    let mut stream = png_writer.stream_writer()?;
    let row_bytes = (img_w as usize) * 4;
    let mut row_buf = vec![0u8; row_bytes];

    let ctx = CellCtx {
        scale,
        cell_w_px: scale,
        cell_h_px: scale,
        grid_w,
        grid_h,
        img_w,
        img_h,
    };

    for gy in 0..grid_h {
        for sy in 0..scale {
            let mut offset = 0usize;
            for gx in 0..grid_w {
                for sx in 0..scale {
                    let rgba = painter(gx, gy, sx, sy, &ctx);
                    row_buf[offset] = rgba[0];
                    row_buf[offset + 1] = rgba[1];
                    row_buf[offset + 2] = rgba[2];
                    row_buf[offset + 3] = rgba[3];
                    offset += 4;
                }
            }
            stream.write_all(&row_buf)?;
        }
    }
    stream.finish()?;
    Ok(())
}

pub fn render_infection_state_png(
    base_dir: impl AsRef<Path>,
    timestep: u32,
    grid_w: u32,
    grid_h: u32,
    healthy_sites: &[(u32, u32)],
    infected_sites: &[(u32, u32)],
    ignored_sites: &[(u32, u32)],
    cfg: &GridRenderConfig,
) -> Result<PathBuf, Box<dyn std::error::Error>> {
    let out_dir = base_dir.as_ref().join("infection");
    let out_path = out_dir.join(format!("{}.png", timestep));

    let mut healthy = vec![false; (grid_w as usize) * (grid_h as usize)];
    let mut infected = vec![false; healthy.len()];
    let mut ignored = vec![false; healthy.len()];

    for &(x, y) in healthy_sites.iter() {
        if x == 0 || y == 0 {
            continue;
        }
        let gx = x - 1;
        let gy = y - 1;
        if gx < grid_w && gy < grid_h {
            healthy[(gy as usize) * (grid_w as usize) + (gx as usize)] = true;
        }
    }
    for &(x, y) in infected_sites.iter() {
        if x == 0 || y == 0 {
            continue;
        }
        let gx = x - 1;
        let gy = y - 1;
        if gx < grid_w && gy < grid_h {
            infected[(gy as usize) * (grid_w as usize) + (gx as usize)] = true;
        }
    }
    for &(x, y) in ignored_sites.iter() {
        if x == 0 || y == 0 {
            continue;
        }
        let gx = x - 1;
        let gy = y - 1;
        if gx < grid_w && gy < grid_h {
            ignored[(gy as usize) * (grid_w as usize) + (gx as usize)] = true;
        }
    }

    let green: [u8; 4] = [0, 128, 0, 255];
    let red: [u8; 4] = [255, 0, 0, 255];
    let blue: [u8; 4] = [0, 0, 255, 255];
    let bg = cfg.background.unwrap_or([0, 0, 0, 0]);

    render_cells_png(&out_path, grid_w, grid_h, cfg, |gx, gy, _ctx| {
        let idx = (gy as usize) * (grid_w as usize) + (gx as usize);
        if infected[idx] {
            red
        } else if healthy[idx] {
            green
        } else if ignored[idx] {
            blue
        } else {
            bg
        }
    })?;

    Ok(out_path)
}

pub fn render_foi_png(
    base_dir: impl AsRef<Path>,
    timestep: u32,
    grid_w: u32,
    grid_h: u32,
    foi: &[f64],
    global_min: f64,
    global_max: f64,
    cfg: &GridRenderConfig,
) -> Result<PathBuf, Box<dyn std::error::Error>> {
    let out_dir = base_dir.as_ref().join("foi");
    let out_path = out_dir.join(format!("{}.png", timestep));

    let use_minmax = global_min.is_finite() && global_max.is_finite() && global_max > global_min;

    render_cells_png(&out_path, grid_w, grid_h, cfg, |gx, gy, _ctx| {
        let idx = (gy as usize) * (grid_w as usize) + (gx as usize);
        let mut v = foi.get(idx).copied().unwrap_or(0.0);
        if !v.is_finite() {
            v = 0.0;
        }
        let mut n = if use_minmax {
            (v - global_min) / (global_max - global_min)
        } else {
            0.0
        };
        if n < 0.0 {
            n = 0.0;
        }
        if n > 1.0 {
            n = 1.0;
        }
        let intensity = (n * 255.0).round() as u8;
        [intensity, intensity, intensity, 255]
    })?;

    Ok(out_path)
}

pub fn render_cells_png_gray16<F>(
    output_path: impl AsRef<Path>,
    grid_w: u32,
    grid_h: u32,
    cfg: &GridRenderConfig,
    mut painter: F,
) -> Result<(), Box<dyn std::error::Error>>
where
    F: FnMut(u32, u32, &CellCtx) -> f64,
{
    let (scale, img_w, img_h) = compute_scale(grid_w, grid_h, cfg)
        .ok_or_else(|| format!("grid too large for max image size"))?;
    if let Some(parent) = output_path.as_ref().parent() {
        fs::create_dir_all(parent)?;
    }

    let file = File::create(&output_path)?;
    let mut writer = BufWriter::new(file);
    let mut png_writer = new_png_encoder_gray16(&mut writer, img_w, img_h)?;
    let mut stream = png_writer.stream_writer()?;
    let row_bytes = (img_w as usize) * 2;
    let mut row_buf = vec![0u8; row_bytes];

    let ctx = CellCtx {
        scale,
        cell_w_px: scale,
        cell_h_px: scale,
        grid_w,
        grid_h,
        img_w,
        img_h,
    };

    for gy in 0..grid_h {
        for _sy in 0..scale {
            let mut offset = 0usize;
            for gx in 0..grid_w {
                let mut n = painter(gx, gy, &ctx);
                if n < 0.0 {
                    n = 0.0;
                }
                if n > 1.0 {
                    n = 1.0;
                }
                let s = (n * 65535.0).round() as u16;
                let be = s.to_be_bytes();
                for _ in 0..scale {
                    row_buf[offset] = be[0];
                    row_buf[offset + 1] = be[1];
                    offset += 2;
                }
            }
            stream.write_all(&row_buf)?;
        }
    }
    stream.finish()?;
    Ok(())
}

pub fn render_foi_png_gray16(
    base_dir: impl AsRef<Path>,
    timestep: u32,
    grid_w: u32,
    grid_h: u32,
    normalized: &[f64],
    cfg: &GridRenderConfig,
) -> Result<PathBuf, Box<dyn std::error::Error>> {
    let out_dir = base_dir.as_ref().join("foi");
    let out_path = out_dir.join(format!("{}.png", timestep));

    render_cells_png_gray16(&out_path, grid_w, grid_h, cfg, |gx, gy, _ctx| {
        let idx = (gy as usize) * (grid_w as usize) + (gx as usize);
        normalized.get(idx).copied().unwrap_or(0.0)
    })?;

    Ok(out_path)
}
