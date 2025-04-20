# Power Trade Green-Rate Prototype Workbook

## Overview

This repo contains a sample Excel workbook (originally developed in Google Sheets) that demonstrates a simple "green‑rate" certification process. The workbook replicates the process of ingesting raw trade data, tagging each trade by energy source, calculating emissions, and visualising key performance indicators in a dashboard.

## Purpose

This prototype illustrates an end‑to‑end workflow for:

1. Importing and tagging trade data by energy source.
2. Looking up emissions factors and certificate classifications.
3. Computing aggregate metrics like green volume, brown volume, and green rate.
4. Calculating total CO₂e and net emissions rate (tCO₂e/MWh).
5. Applying threshold‑based flags.
6. Building a one‑page dashboard with charts and status indicators.

The focus is on creating a transparent model in a lightweight format.

## Workbook Structure

The workbook includes five sheets:

| Sheet Name                    | Description                                                                               |
| ----------------------------- | ----------------------------------------------------------------------------------------- |
| **Trades**                    | Raw trade records (TradeID, Volume (MWh), Source, Type)                                   |
| **Emissions_Factor_Lookup**   | Emissions factor table mapping Source to EF (tCO₂e/MWh)                                   |
| **Certification_Lookup**      | Certificate lookup table (CertificateID, CertType, Source → Classification)               |
| **Calc**                      | Core calculations: volumes, rates, emissions aggregation, and status flags                |
| **Dashboard**                 | One‑page summary of key metrics and visualizations                                        |

## Named Ranges and Columns

To simplify formulas, the following named ranges are defined:

| Named Range                  | Range Address                    | Description                                                |
| ---------------------------- | -------------------------------- | ---------------------------------------------------------- |
| **Trades_Raw**              | Trades!A2:D100                   | Raw trade data (TradeID, Volume, Source, Type)             |
| **Emissions_Factor_Table**  | Emissions_Factor_Lookup!A2:B7    | Emissions factor lookup by Source                          |
| **Cert_Table**              | Certification_Lookup!A2:D7       | Certificate table for Source → Classification               |
| **Volume_Col**              | Trades!B2:B100                   | Volume column in Trades                                    |
| **Type_Col**                | Trades!D2:D100                   | Type (Green/Brown) column in Trades                        |
| **CO2e_Col**                | Trades!F2:F100                   | Computed CO₂e per trade (Volume × EF)                      |

## How It Works

1. **Tagging** (Trades sheet)  
   - Each trade’s `Source` is classified into `Green` or `Brown` via a lookup in `Cert_Table`.  
2. **Lookup Emissions Factors** (Trades sheet)  
   - `VLOOKUP(Source, Emissions_Factor_Table, 2)` populates EF in tCO₂e/MWh.  
3. **Calculate CO₂e** (Trades sheet)  
   - `CO2e = Volume × EF` for each trade.  
4. **Aggregate Metrics** (Calc sheet)  
   - **Total Volume** = `SUM(Volume_Col)`  
   - **Green Volume** = `SUMIF(Type_Col, "Green", Volume_Col)`  
   - **Brown Volume** = Total − Green  
   - **Green Rate** = Green Volume / Total Volume  
   - **Total CO₂e** = `SUM(CO2e_Col)`  
   - **Net Emissions Rate** = Total CO₂e / Total Volume  
5. **Threshold Flags** (Calc sheet)  
   - Compare Green Rate and Net Emissions Rate against preset targets using `IF` statements and conditional formatting.  
6. **Dashboard Visualisation** (Dashboard sheet)  
   - Pull in key KPIs and status flags.  
   - A pie chart for green vs. brown volume.  
   - A bar chart for net emissions rate.
