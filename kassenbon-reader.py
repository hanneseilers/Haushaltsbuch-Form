from sys import argv
import functions as f
import config as cfg


def process(subdir:str) -> None:
    for filepath in f.get_files_from_dir(subdir):
        bon = f.read_bon_data_from_file(
            filepath=filepath,
            data_seperator=cfg.bon_data_seperator)

        print(bon)


# ---- MAIN ----
if __name__ == '__main__':
    bon_dir = cfg.bon_dir_default
    if len(argv) > 1:
        bon_dir = argv[1].strip()

    process(bon_dir)