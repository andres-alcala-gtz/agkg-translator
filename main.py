import copy
import pathlib

import move
import translate
import utilities


if __name__ == "__main__":

    processes = 8

    directory_src = pathlib.Path("input")
    directory_dst = pathlib.Path("output")

    watch = utilities.Watch()

    move.move(directory_src, directory_dst, copy.deepcopy(watch))
    translate.translate(directory_dst, processes, copy.deepcopy(watch))
    translate.translate(directory_dst, 1, copy.deepcopy(watch))
