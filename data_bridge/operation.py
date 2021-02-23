from data_bridge.data import (
    get_master_data,
    place_data_into_new_master_format,
    root_path,
    Master,
    overall_dashboard,
    get_project_information,
    DandelionData, make_a_dandelion_auto,
)


# wb = place_data_into_new_master_format(data[0])
# wb.save(root_path/"output/CDG_new_data.xlsx")

# d = get_master_data()
m = Master(get_master_data(), get_project_information())
# db_m = root_path / "input/cdg_dashboard_master.xlsx"   # dashboard master
# db = overall_dashboard(m, db_m)
# db.save(root_path / "output/cdg_dashboard_compiled.xlsx")
dan_l = DandelionData(m)
make_a_dandelion_auto(dan_l)
