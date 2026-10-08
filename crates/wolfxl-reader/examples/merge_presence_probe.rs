//! Reader-only diagnostic: exact XML merge scan versus conservative byte probe.
//! This does not measure Python assignment or edit/save throughput.
use std::time::Instant;
use wolfxl_reader::NativeXlsxBook;

fn median(samples: &[f64]) -> f64 {
    let mut sorted = samples.to_vec();
    sorted.sort_by(f64::total_cmp);
    sorted[sorted.len() / 2]
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let path = args
        .get(1)
        .expect("usage: merge_presence_probe PATH SHEET [ITERATIONS]");
    let sheet = args.get(2).expect("missing sheet title");
    let iterations: usize = args
        .get(3)
        .map_or(7, |s| s.parse().expect("invalid iterations"));
    assert!(iterations > 0);
    let book = NativeXlsxBook::open_path(path).expect("open source");
    assert!(book
        .worksheet_merged_ranges(sheet)
        .expect("exact warmup")
        .is_empty());
    assert!(!book
        .worksheet_may_have_merged_cells(sheet)
        .expect("probe warmup"));
    let mut exact = Vec::new();
    let mut probe = Vec::new();
    for _ in 0..iterations {
        let started = Instant::now();
        let refs = book.worksheet_merged_ranges(sheet).expect("exact scan");
        exact.push(started.elapsed().as_secs_f64());
        assert!(refs.is_empty());
        let started = Instant::now();
        let present = book
            .worksheet_may_have_merged_cells(sheet)
            .expect("byte probe");
        probe.push(started.elapsed().as_secs_f64());
        assert!(!present);
    }
    println!(
        "{{\"scope\":\"reader-only source scan; not assignment/edit-save throughput\",\"fixture\":{path:?},\"sheet\":{sheet:?},\"exact_samples_seconds\":{exact:?},\"probe_samples_seconds\":{probe:?},\"exact_median_seconds\":{},\"probe_median_seconds\":{},\"speedup\":{}}}",
        median(&exact), median(&probe), median(&exact) / median(&probe)
    );
}
