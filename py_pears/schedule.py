import os
import py_pears.utils as utils
import py_pears.reports.sites_report as sites_report


# Calculate the path to the root directory of this package
ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '.'))

# Set directories using package defaults
EXPORT_DIR = ROOT_DIR + '/pears_exports/'
OUT_DIR = ROOT_DIR + "/reports/outputs/"

creds = utils.load_credentials()

# Run Sites Report with default inputs
sites_report.main(creds=creds, export_dir=EXPORT_DIR, output_dir=OUT_DIR)
