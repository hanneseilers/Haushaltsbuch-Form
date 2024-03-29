# ---- CONFIGURATION FILE ----

# ---- HAUSHALTSBUCH READER ----
# file names
base_filename = "Haushaltsbuch.xlsm"
monthly_filename = "%y_%d_" + base_filename
backup_filename = monthly_filename.replace(".xlsm", "_backup.xlsm")

# config for header data
header_row = 1
header_column_start = 3     # column = C
header_column_end = 40      # column = AN

# config of date data
date_column = 1             # column = A
date_row_start = 3
date_row_end = 33

# ---- KASSENBON READER ----
bon_dir_default = "bons"
bon_data_seperator = ';'
bon_content_key_name = "article"
bon_content_key_value = "current"
