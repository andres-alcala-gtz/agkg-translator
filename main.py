from time import perf_counter
from datetime import timedelta

from move import move_mso
from update import update_mso
from translate import translate_mso
from outdate import outdate_mso


if __name__ == "__main__":

    time_beginning = perf_counter()

    move_mso()
    update_mso()
    translate_mso()
    outdate_mso()

    print(f"COMPLETED IN {timedelta(seconds=int(perf_counter() - time_beginning))}")
