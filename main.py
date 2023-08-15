from time import perf_counter
from datetime import timedelta

from move import move
from update import update
from translate import translate
from outdate import outdate


if __name__ == "__main__":

    time_beginning = perf_counter()

    move()
    update()
    translate()
    outdate()

    print(f"COMPLETED IN {timedelta(seconds=int(perf_counter() - time_beginning))}")
