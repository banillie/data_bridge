from data_bridge.data import (
    get_master_data,
    place_data_into_new_master_format,
    root_path,
    Master,
    overall_dashboard,
    get_project_information,
    DandelionData,
    run_dandelion_matplotlib_chart,
    dandelion_data_into_wb,
    DandelionChart
)


# wb = place_data_into_new_master_format(data[0])
# wb.save(root_path/"output/CDG_new_data.xlsx")

# d = get_master_data()
m = Master(get_master_data(), get_project_information())
# db_m = root_path / "input/cdg_dashboard_master.xlsx"   # dashboard master
# db = overall_dashboard(m, db_m)
# db.save(root_path / "output/cdg_dashboard_compiled.xlsx")
dan = DandelionData(m, quarter=["Q3 20/21"], group=["CF"])
run_dandelion_matplotlib_chart(dan.d_data["Q3 20/21"], chart=True)
