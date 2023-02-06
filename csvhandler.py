import os

def get_csvs_from_input_folder(desired_path = "input/"):
    
    if (os.path.exists(desired_path)):
        all_files = os.listdir(desired_path)
        csv_files = list(filter(lambda f: f.endswith('.csv'), all_files))
        csv_files = [desired_path + file for file in csv_files]
        return csv_files

    return