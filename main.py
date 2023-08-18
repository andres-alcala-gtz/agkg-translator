from copy import deepcopy
from time import perf_counter

from move import move
from translate import translate
from utilities import Watch


if __name__ == "__main__":

    watch = Watch(time_beginning=perf_counter())

    move(watch=deepcopy(watch))
    translate(watch=deepcopy(watch))
    translate(watch=deepcopy(watch))
