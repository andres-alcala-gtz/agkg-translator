import copy
import pathlib

import move
import translate
import utilities


if __name__ == "__main__":

    processes = 8
    full = False
    safe = True

    directory_src = pathlib.Path("input")
    directory_dst = pathlib.Path("output")

    watch = utilities.Watch()

    move.move(full, directory_src, directory_dst, copy.deepcopy(watch))
    translate.translate(safe, directory_dst, processes, copy.deepcopy(watch))
    translate.translate(safe, directory_dst, 1, copy.deepcopy(watch))
