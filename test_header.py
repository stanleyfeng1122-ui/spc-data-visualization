"""Quick test to verify header band positioning."""
import sys
sys.path.insert(0, "/Users/zhefeng/Desktop/Vibe Coding/Data Visualiztion")

from spc_parser import parse_excel
from chart_utils import prepare_combined_data, build_combined_chart, finalize_plotly_style
import json

files_to_load = [
    "/Users/zhefeng/Desktop/Vibe Coding/Data Visualiztion/FX_K116_P1_PP_100%.xlsx",
    "/Users/zhefeng/Desktop/Vibe Coding/Data Visualiztion/TRM_X3083_BC,P2_PP_100% Data_0127.xlsx",
    "/Users/zhefeng/Desktop/Vibe Coding/Data Visualiztion/TY_K116_P1_PP_100%.xlsx",
]

parsed_files = []
for f in files_to_load:
    try:
        result = parse_excel(f)
        parsed_files.append({
            "filename": f.split("/")[-1],
            "factory": result.factory,
            "part_number": result.part_number,
            "dimensions": result.dimensions,
            "data": result.data,
            "meta_columns": result.meta_columns,
        })
    except Exception as e:
        print(f"Error parsing {f}: {e}")

if not parsed_files:
    print("No files parsed")
    sys.exit(1)

# Get first dimension
all_dims = []
for pf in parsed_files:
    for d in pf["dimensions"]:
        if d not in all_dims:
            all_dims.append(d)

print(f"Factories: {[pf['factory'] for pf in parsed_files]}")
print(f"Dimensions: {all_dims[:3]}")

dim_nos = all_dims[:1]  # Just first dim
df, dim_metas = prepare_combined_data(parsed_files, dim_nos)

if df is not None:
    fig = build_combined_chart(
        df, dim_metas, dim_nos,
        section_by_fields=["Factory"],
        color_by="Raw material",
        y_axis_mode="Deviation from Nominal",
        exclude_intervals=False,
        group_label="Test",
    )
    if fig:
        fig = finalize_plotly_style(fig)
        # Extract header shapes and annotations for verification
        shapes = fig.layout.shapes
        annotations = fig.layout.annotations

        print(f"\n--- Header Shapes ({len(shapes)}) ---")
        for i, s in enumerate(shapes):
            print(f"  Shape {i}: x0={s.x0:.4f}, x1={s.x1:.4f}, y0={s.y0}, y1={s.y1}, xref={s.xref}")

        print(f"\n--- Annotations ({len(annotations)}) ---")
        for i, a in enumerate(annotations):
            if a.yref == "paper" and a.y and a.y > 0.9:
                print(f"  Ann {i}: text='{a.text}', x={a.x:.4f}, y={a.y}, xref={a.xref}, xanchor={a.xanchor}")

        # Save as HTML for visual inspection
        fig.write_html("/tmp/spc_test_chart.html")
        print("\nChart saved to /tmp/spc_test_chart.html")

        # Also save as image if kaleido available
        try:
            fig.write_image("/tmp/spc_test_chart.png", width=1400, height=700)
            print("Image saved to /tmp/spc_test_chart.png")
        except Exception as e:
            print(f"Could not save image: {e}")
    else:
        print("Chart build returned None")
else:
    print("No data prepared")
