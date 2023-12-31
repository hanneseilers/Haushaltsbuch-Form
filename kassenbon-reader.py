from sys import argv
import functions as f
import config as cfg


def process(subdir:str) -> None:
    for filepath in f.get_files_from_dir(subdir):
        bon_data = f.read_bon_data_from_file(
            filepath=filepath,
            data_seperator=cfg.bon_data_seperator,
            data_name_key=cfg.bon_content_key_name,
            data_value_key=cfg.bon_content_key_value)

        if not bon_data:
            continue

        print(bon_data)
        bon_filepath, bon_date, bon_content = bon_data

        # process each article
        for _data in bon_content:
            if cfg.bon_content_key_name in _data.keys() and cfg.bon_content_key_value in _data.keys():

                _article_name = (str(_data[cfg.bon_content_key_name])
                                 .strip())
                _current_value = (str(_data[cfg.bon_content_key_value])
                                  .replace("€", '')
                                  .replace("EUR", '')
                                  .replace(',', '.')
                                  .strip())

                print(f"{_article_name}: {_current_value} €")


# ---- MAIN ----
if __name__ == '__main__':
    bon_dir = cfg.bon_dir_default
    if len(argv) > 1:
        bon_dir = argv[1].strip()

    process(bon_dir)