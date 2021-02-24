from analysis_engine.data import open_word_doc, put_matplotlib_fig_into_word

from data_bridge.data import (
    get_master_data,
    place_data_into_new_master_format,
    root_path,
    Master,
    overall_dashboard,
    get_project_information,
    DandelionData,
    make_a_dandelion_auto_cdg,
    convert_pdf_to_png,
    run_p_reports_cdg,
)


# wb = place_data_into_new_master_format(data[0])
# wb.save(root_path/"output/CDG_new_data.xlsx")

# d = get_master_data()
m = Master(get_master_data(), get_project_information())
## dashboard master
# db_m = root_path / "input/cdg_dashboard_master.xlsx"
# db = overall_dashboard(m, db_m)
# db.save(root_path / "output/cdg_dashboard_compiled.xlsx")
## dandelion
# dl_data = DandelionData(m)
# d_lion = make_a_dandelion_auto_cdg(dl_data)
# doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
# put_matplotlib_fig_into_word(doc, d_lion, size=7.5)
# doc.save(root_path / "output/dandelion_graph.docx")
## summaries
run_p_reports_cdg(m)