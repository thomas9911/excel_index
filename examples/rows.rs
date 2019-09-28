use excel_index::ExcelIndex;

fn main() {
    let e1: ExcelIndex = "A4".parse().unwrap();

    for cell in e1.into_iter().take(15) {
        println!("{}", cell)
    }

    let e1: ExcelIndex = "A4".parse().unwrap();
    let e2: ExcelIndex = "G14".parse().unwrap();

    for cell in e1.into_range(e2) {
        println!("{}", cell)
    }
}
