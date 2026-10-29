#!/usr/bin/env bash

IMAGE="translator"

case "$1" in

    build)
        docker build -t "$IMAGE" .
        ;;

    run)
        docker run -it --rm -v "$(pwd):/app" "$IMAGE"
        ;;

    remove)
        docker image rm "$IMAGE"
        ;;

    dev)
        docker run -it --rm -v "$(pwd):/app" "$IMAGE" bash
        ;;

    *)
        echo "Usage: bash docker.sh {build|run|remove|dev}"
        exit 1
        ;;

esac
