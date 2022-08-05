import mk32_main as main


csm = main.FilePath(
    main.path_dict['customer_center_folder_path'], main.todays_date, 'csm')

print(csm.__str__())
