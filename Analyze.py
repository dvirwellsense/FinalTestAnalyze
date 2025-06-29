import pandas as pd
import numpy as np
from pathlib import Path
from scipy.ndimage import label, center_of_mass
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io

def extract_third_matrix(file_path):
    df = pd.read_csv(file_path, header=None)
    row_indices = df[df[0].astype(str).str.startswith("row_0")].index.tolist()
    if len(row_indices) < 3:
        raise ValueError(f"Less than 3 matrices found in {file_path}")
    start_idx = row_indices[2]
    df_matrix = df.iloc[start_idx:, 1:]
    df_matrix = df_matrix.dropna(axis=1, how='all')
    return df_matrix.astype(float).to_numpy()

def find_weight_spots(diff_matrix, threshold_ratio=0.1, num_weights=3):
    max_val = np.max(diff_matrix)
    threshold = threshold_ratio * max_val
    binary = (diff_matrix > threshold).astype(int)
    labeled, num_features = label(binary)
    spots = center_of_mass(diff_matrix, labeled, range(1, num_features + 1))
    spot_values = [diff_matrix[int(y)][int(x)] for y, x in spots]
    sorted_spots = sorted(zip(spot_values, spots), reverse=True)[:num_weights]
    return [(int(y), int(x), val) for val, (y, x) in sorted_spots], labeled

def get_stats_for_label(diff_matrix, labeled_matrix, label_id):
    mask = labeled_matrix == label_id
    values = diff_matrix[mask]
    return np.mean(values), np.max(values)

def classify_files_by_avg(folder_path):
    csv_files = list(Path(folder_path).rglob("*.csv"))
    if len(csv_files) != 2:
        raise ValueError(f"Expected exactly 2 CSV files in {folder_path}, found {len(csv_files)}")
    matrices = [extract_third_matrix(f) for f in csv_files]
    # Ignore inactive rows and cols when computing averages for classification
    avgs = [np.mean(mat[1:-1, 1:]) for mat in matrices]
    if avgs[0] < avgs[1]:
        return csv_files[0], csv_files[1], avgs[0]
    else:
        return csv_files[1], csv_files[0], avgs[1]

def show_debug_plot(mat_diff, spots, folder_name):
    plt.figure(figsize=(8,6))
    plt.imshow(mat_diff, cmap='hot', interpolation='nearest')
    ys, xs, _ = zip(*spots)
    plt.scatter(xs, ys, marker='x', color='cyan', s=150, label='Detected weights')
    plt.title(f"Weight spots in {folder_name}")
    plt.colorbar(label="Difference")
    plt.legend()
    plt.show()

def analyze_sheet(folder_path, debug=False):
    empty_file, weight_file, empty_avg_clean = classify_files_by_avg(folder_path)
    mat_empty = extract_third_matrix(empty_file)[1:-1, 1:]
    mat_weight = extract_third_matrix(weight_file)[1:-1, 1:]
    mat_diff = mat_weight - mat_empty

    mean_diff = np.mean(mat_diff)
    std_diff = np.std(mat_diff)
    total_diff = np.sum(mat_diff)

    spots, labeled = find_weight_spots(mat_diff)
    sorted_spots = sorted(spots, key=lambda p: p[0])  # sort by row (y)

    weights_lbs = [20, 10, 5]
    weight_responses = {}

    for w, (y, x, _) in zip(weights_lbs, sorted_spots):
        label_id = labeled[y, x]
        mean_val, max_val = get_stats_for_label(mat_diff, labeled, label_id)
        weight_responses[f"{w}lb_response"] = mean_val
        weight_responses[f"{w}lb_max"] = max_val

    if debug:
        print(f"Debug info for {Path(folder_path).name}:")
        for w in weights_lbs:
            print(f"  {w}lb mean: {weight_responses[f'{w}lb_response']:.4e}, max: {weight_responses[f'{w}lb_max']:.4e}")
        show_debug_plot(mat_diff, spots, Path(folder_path).name)

    return {
        "Mat": Path(folder_path).name,
        "mean_diff": mean_diff,
        "std_diff": std_diff,
        "total_diff": total_diff,
        "empty_avg_clean": empty_avg_clean,
        **weight_responses,
        "ratio_10_to_20": weight_responses["10lb_response"] / weight_responses["20lb_response"],
        "ratio_5_to_20": weight_responses["5lb_response"] / weight_responses["20lb_response"],
        "weight_file": weight_file.name
    }

def add_mat_slide(prs, row, base_folder):
    layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)

    # Add table with data (6 rows, 2 cols)
    rows, cols = 6, 2
    table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.0), Inches(5.0), Inches(2.5)).table

    # Set column widths
    table.columns[0].width = Inches(2)
    table.columns[1].width = Inches(3)

    # Fill table cells
    table.cell(0, 0).text = "Baseline avg (no weight)"
    table.cell(0, 1).text = f"{row['empty_avg_clean']:.2e}"

    table.cell(1, 0).text = "20lb: mean / max"
    table.cell(1, 1).text = f"{row['20lb_response']:.2e} / {row['20lb_max']:.2e}"

    table.cell(2, 0).text = "10lb: mean / max"
    table.cell(2, 1).text = f"{row['10lb_response']:.2e} / {row['10lb_max']:.2e}"

    table.cell(3, 0).text = "5lb: mean / max"
    table.cell(3, 1).text = f"{row['5lb_response']:.2e} / {row['5lb_max']:.2e}"

    table.cell(4, 0).text = "Ratio 10lb / 20lb"
    table.cell(4, 1).text = f"{row['ratio_10_to_20']:.2f}"

    table.cell(5, 0).text = "Ratio 5lb / 20lb"
    table.cell(5, 1).text = f"{row['ratio_5_to_20']:.2f}"

    # Add image on the right, aligned to bottom
    folder = Path(base_folder) / row["Mat"]
    image_file = folder / row["weight_file"].replace("_rawData.csv", "_heatmap.png")
    if image_file.exists():
        pic = slide.shapes.add_picture(str(image_file), Inches(6), Inches(0), width=Inches(3))
        pic.top = prs.slide_height - pic.height
    else:
        print(f"Warning: Image not found: {image_file}")

    # Add title textbox centered at bottom
    textbox = slide.shapes.add_textbox(Inches(0.5), prs.slide_height - Inches(0.7), Inches(8), Inches(0.5))
    tf = textbox.text_frame
    p = tf.add_paragraph()
    p.text = f"Mat {row['Mat']}"
    p.font.size = Pt(24)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER

def add_summary_slide(prs, df):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Summary Table"
    rows, cols = len(df) + 1, 7
    table = slide.shapes.add_table(rows, cols, Inches(0.2), Inches(1), Inches(9), Inches(0.3 * rows)).table
    headers = ["Mat", "20lb", "10lb", "5lb", "10/20", "5/20", "Baseline"]
    for col, h in enumerate(headers):
        table.cell(0, col).text = h

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        table.cell(i, 0).text = str(row["Mat"])
        table.cell(i, 1).text = f"{row['20lb_response']:.2e}"
        table.cell(i, 2).text = f"{row['10lb_response']:.2e}"
        table.cell(i, 3).text = f"{row['5lb_response']:.2e}"
        table.cell(i, 4).text = f"{row['ratio_10_to_20']:.2f}"
        table.cell(i, 5).text = f"{row['ratio_5_to_20']:.2f}"
        table.cell(i, 6).text = f"{row['empty_avg_clean']:.2e}"

def add_plot_slide(prs, df):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Final Test Summary Graphs"

    fig, ax = plt.subplots(figsize=(10, 6), dpi=300)
    for w in ["5lb_response", "10lb_response", "20lb_response"]:
        ax.plot(df["Mat"], df[w], label=w)
    ax.set_xlabel("Mat")
    ax.set_ylabel("Mean Response")
    ax.legend()
    ax.grid(True)

    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', dpi=300, bbox_inches='tight')
    plt.close(fig)
    img_stream.seek(0)

    slide.shapes.add_picture(img_stream, Inches(0.5), Inches(1), width=Inches(9))

def create_ppt_report(df, base_folder):
    prs = Presentation()

    # Add one slide per mat
    for _, row in df.iterrows():
        add_mat_slide(prs, row, base_folder)

    # Add summary slide
    add_summary_slide(prs, df)

    # Add plot slide
    add_plot_slide(prs, df)

    output_path = Path(base_folder) / "final_test_analysis_report.pptx"
    prs.save(output_path)
    print(f"PowerPoint report saved to {output_path}")

def analyze_all(base_folder, debug=False):
    results = []
    for subfolder in Path(base_folder).iterdir():
        if subfolder.is_dir():
            try:
                print(f"Analyzing {subfolder.name}...")
                result = analyze_sheet(subfolder, debug=debug)
                results.append(result)
            except Exception as e:
                print(f"Error in {subfolder.name}: {e}")

    df = pd.DataFrame(results)
    xlsx_path = Path(base_folder) / "final_test_analysis_summary.xlsx"
    df.to_excel(xlsx_path, index=False)
    print(f"Excel saved to {xlsx_path}")

    create_ppt_report(df, base_folder)

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Analyze final test data from calibration mats.")
    parser.add_argument("folder", help="Path to folder containing mats subfolders")
    parser.add_argument("--debug", action="store_true", help="Enable debug output and visualization")
    args = parser.parse_args()
    analyze_all(args.folder, debug=args.debug)
