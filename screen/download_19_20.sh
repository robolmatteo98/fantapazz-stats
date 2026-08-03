#!/bin/sh
set -e

node screen/download_campionato.js --anno 19_20 --giornata 495 --lega 145707
node screen/download_fantapunti.js --anno 19_20 --giornata 495 --lega 145707

node screen/download_champions.js --anno 19_20 --giornata 521 --lega 145707 --fase semifinali_andata
node screen/download_champions.js --anno 19_20 --giornata 524 --lega 145707 --fase semifinali_ritorno
node screen/download_champions.js --anno 19_20 --giornata 528 --lega 145707 --fase finale

node screen/download_europa_league.js --anno 19_20 --giornata 521 --lega 145707 --fase semifinali_andata
node screen/download_europa_league.js --anno 19_20 --giornata 524 --lega 145707 --fase semifinali_ritorno
node screen/download_europa_league.js --anno 19_20 --giornata 528 --lega 145707 --fase finale

node screen/download_coppa_italia.js --anno 19_20 --giornata 506 --lega 145707 --fase quarti_andata
node screen/download_coppa_italia.js --anno 19_20 --giornata 518 --lega 145707 --fase quarti_ritorno
node screen/download_coppa_italia.js --anno 19_20 --giornata 522 --lega 145707 --fase semifinali_andata
node screen/download_coppa_italia.js --anno 19_20 --giornata 525 --lega 145707 --fase semifinali_ritorno
node screen/download_coppa_italia.js --anno 19_20 --giornata 529 --lega 145707 --fase finale
