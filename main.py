import time
import copy

import move
import translate
import utilities


if __name__ == "__main__":

    watch = utilities.Watch(time_beginning=time.perf_counter())

    move.move(watch=copy.deepcopy(watch))
    translate.translate(watch=copy.deepcopy(watch))
    translate.translate(watch=copy.deepcopy(watch))
