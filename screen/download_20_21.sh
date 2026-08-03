#!/bin/sh
set -e

node screen/download_campionato.js --anno 20_21 --giornata 533 --lega 201910
node screen/download_fantapunti.js --anno 20_21 --giornata 533 --lega 201910

node screen/download_champions.js --anno 20_21 --giornata 563 --lega 201910 --fase semifinali_andata
node screen/download_champions.js --anno 20_21 --giornata 564 --lega 201910 --fase semifinali_ritorno
node screen/download_champions.js --anno 20_21 --fase finale --url "https://www.fantapazz.com/fantacalcio/formazioni-in-campo-lega/567/5/201910"

node screen/download_europa_league.js --anno 20_21 --giornata 563 --lega 201910 --fase semifinali_andata

node screen/download_coppa_italia.js --anno 20_21 --giornata 547 --lega 201910 --fase quarti_andata
node screen/download_coppa_italia.js --anno 20_21 --giornata 548 --lega 201910 --fase quarti_ritorno
node screen/download_coppa_italia.js --anno 20_21 --giornata 560 --lega 201910 --fase semifinali_andata
node screen/download_coppa_italia.js --anno 20_21 --giornata 561 --lega 201910 --fase semifinali_ritorno
node screen/download_coppa_italia.js --anno 20_21 --giornata 566 --lega 201910 --fase finale
