from sys import argv
import functions as f
import config as cfg


def process(bon_dir:str) -> None:
    print(f.get_files_from_dir(bon_dir))


# ---- MAIN ----
if __name__ == '__main__':
    bon_dir = cfg.bon_dir_default
    if len(argv) > 1:
        bon_dir = argv[1].strip()

    process(bon_dir)