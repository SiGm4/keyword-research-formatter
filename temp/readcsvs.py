import glob

def configure_csv_files():
    csv_files = {}
    csv_files["Meta Descriptions Too Long"] = ""
    csv_files["Meta Descriptions Too Short"] = ""
    csv_files["Meta Descriptions Tag Missing"] = ""
        
    mylist = [f for f in glob.glob("*.csv")]
    
    for file in mylist:
        if "meta-description-too-long" in file:
            csv_files["Meta Descriptions Too Long"] = file
        elif "meta-description-too-short" in file:
            csv_files["Meta Descriptions Too Short"] = file
        elif "meta-description-tag-miss" in file:
            csv_files["Meta Descriptions Tag Missing"] = file


    for key in csv_files:
        if csv_files[key] == "":
            csv_files[key] = input("Give filename for " + key + ", including the extension: ") #TODO add a "or none to ignore" and implement functionality

    for key in csv_files:
        print(key + ":", csv_files[key])
    
    return csv_files
