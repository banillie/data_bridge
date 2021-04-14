from analysis_engine.data import open_word_doc, put_matplotlib_fig_into_word

from data_bridge.data import (
    get_master_data,
    place_data_into_new_master_format,
    root_path,
    Master,
    overall_dashboard,
    get_project_information,
    DandelionData,
    make_a_dandelion_auto,
)


# wb = place_data_into_new_master_format(data[0])
# wb.save(root_path/"output/CDG_new_data.xlsx")

# d = get_master_data()
m = Master(get_master_data(), get_project_information())
## dashboard master
# db_m = root_path / "input/cdg_dashboard_master.xlsx"
# db = overall_dashboard(m, db_m)
# db.save(root_path / "output/cdg_dashboard_compiled.xlsx")

CDG_DIR = ["CFPD", "GF", "SCS"]
# CDG_DIR = ["SCS"]

# dandelion
op_args = {
    "quarter": ["Q4 20/21"],
    "group": CDG_DIR,
    "chart": True,
    }
dl_data = DandelionData(m, **op_args)
d_lion = make_a_dandelion_auto(dl_data)
doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
put_matplotlib_fig_into_word(doc, d_lion, size=7.5)
doc.save(root_path / "output/cdg_dandelion_graph.docx")

# ## summaries
# run_p_reports_cdg(m)