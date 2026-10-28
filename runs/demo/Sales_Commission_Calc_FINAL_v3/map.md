# Map: Sales_Commission_Calc_FINAL_v3__USE_THIS_.xlsm

Outputs = formula cells on sheet(s) ['Summary']. 43 output cells, 468 true input cells in 23 blocks, 1161 formulas, 12 dynamic references, 0 circular group(s).

Each line is a column block. `!` lines are cells that break the block's pattern. `<~ dynamic` is an edge the static map cannot confirm (OFFSET/INDIRECT).

## Output chains

### Summary!C5:C16

```
Summary!C5:C16 =SUMIF(Calc!$B$2:$B$61,$A5,Calc!$S$2:$S$61)+SUMIF(adj!$A$2:$A$20,$A5,adj!$B$2:$B$20)  (x12)
  <- Calc!B2:B61 =Inputs!B2  (x60)
    <- Inputs!B2:B61 typed values (60) e.g. ["Sarah O'Connor", 'Rohan Mehta ', 'Wei Zhang']
  <- Calc!S2:S61 =N2  (x38)
        ! 22 typed value(s) inside this formula run, rows 40, 41, 42, 43, 44, 45, 46, 47 ... (+14)
    <- Calc!N2:N61 =L2+M2  (x60)
      <- Calc!L2:L61 =F3*K3*IFERROR(VLOOKUP(B3,Rates!$J$2:$K$13,2,FALSE),0)  (x45)
            ! 15 cell(s) use a different formula, rows 2, 23, 27, 31, 34, 38, 39, 41 ... (+7): =F2*K2*IFERROR(VLOOKUP(B2,Rates!$J$2:$K$13,2,FALSE),0)*1.07
        <- Calc!B2:B61 (expanded above)
        <- Calc!F2:F61 =N(Inputs!E2)  (x60)
          <- Inputs!E2:E61 typed values (60) e.g. [190000, 209000, 250000]
        <- Calc!K2:K61 =MIN(H2+I2+J2,0.18)  (x60)
          <- Calc!H2:H61 =VLOOKUP(F2,Rates!$A$2:$B$10,2,TRUE)  (x60)
            <- Calc!F2:F61 (expanded above)
            <- Rates!A2:A10 typed values (9) e.g. [0, 25000, 50000]
            <- Rates!B2:B10 typed values (9) e.g. [0.02, 0.03, 0.04]
          <- Calc!I2:I61 =INDEX(Rates!$E$2:$E$10,G2)  (x60)
            <- Calc!G2:G61 =IF(F2>=1000000,9,IF(F2>=600000,8,IF(F2>=400000,7,IF(AND(C2="EMEA",F2>=250000),6,IF(F2>250000,6,IF(F2>=150000,5,IF(F2>=100000,4,IF(F2>=50000,3,IF(F2>=25000,2,1)))))))))  (x60)
              <- Calc!C2:C61 =Inputs!C2  (x60)
                <- Inputs!C2:C61 typed values (60) e.g. ['EMEA', 'India', 'APAC']
              <- Calc!F2:F61 (expanded above)
            <- Rates!E2:E10 typed values (9) e.g. [0, 0, 0.0025]
          <- Calc!J2:J61 =IFERROR(VLOOKUP(D2,Rates!$G$2:$H$6,2,FALSE),0)  (x60)
            <- Calc!D2:D61 =Inputs!D2  (x60)
              <- Inputs!D2:D61 typed values (60) e.g. ['Support Plus', 'Core Platform', 'Support Plus']
            <- Rates!G2:G6 typed values (5) e.g. ['Core Platform', 'Analytics Add-on', 'Enterprise Suite']
            <- Rates!H2:H6 typed values (5) e.g. [0, 0.01, 0.04]
        <- Rates!J2:J13 typed values (12) e.g. ['Jessica Alvarez', 'Michael Brooks', 'Daniel Kim']
        <- Rates!K2:K13 typed values (12) e.g. [1, 1, 0.5]
      <- Calc!M2:M61 =IF(Inputs!H2>90,-0.02*L2,0)  (x60)
        <- Calc!L2:L61 (expanded above)
        <- Inputs!H2:H61 typed values (60) e.g. [85, 32, 104]
  <- Summary!A5:A16 typed values (12) e.g. ['Jessica Alvarez', 'Michael Brooks', 'Daniel Kim']
  <- adj!A2:A3 typed values (2) e.g. ["Sarah O'Connor", 'Wei Zhang'] [veryHidden sheet]
  <- adj!B2:B3 typed values (2) e.g. [4500, -1200] [veryHidden sheet]
```

### Summary!C19:C22

```
Summary!C19:C22 =SUMIF(Calc!$C$2:$C$61,$A19,Calc!$S$2:$S$61)  (x4)
  <- Calc!C2:C61 (expanded above)
  <- Calc!S2:S61 (expanded above)
  <- Summary!A19:A22 typed values (4) e.g. ['AMER', 'EMEA', 'APAC']
```

### Summary!C24:C26

```
Summary!C24:C26 =SUM(C5:C16)  (x1)
      ! 1 cell(s) use a different formula, rows 25: =SUM(Inputs!E2:E61)
      ! 1 cell(s) use a different formula, rows 26: =C24/C25
  <- Inputs!E2:E61 (expanded above)
  <- Summary!C5:C16 =SUMIF(Calc!$B$2:$B$61,$A5,Calc!$S$2:$S$61)+SUMIF(adj!$A$2:$A$20,$A5,adj!$B$2:$B$20)  (x12)
    <- Calc!B2:B61 (expanded above)
    <- Calc!S2:S61 (expanded above)
    <- Summary!A5:A16 (expanded above)
    <- adj!A2:A3 (expanded above)
    <- adj!B2:B3 (expanded above)
```

### Summary!D5:D16

```
Summary!D5:D16 =SUMIF(Calc!$B$2:$B$61,$A5,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1))  (x12)
  <- Calc!B2:B61 (expanded above)
  <- Calc!O1:O61 =IF(ROUNDUP(MONTH($E2)/3,0)=1,$N2,0)  (x60)
        ! 1 typed value(s) inside this formula run, rows 1  <~ dynamic-base
    <- Calc!E2:E61 =Inputs!G2  (x60)
      <- Inputs!G2:G61 typed values (60) e.g. [datetime.datetime(2026, 1, 1, 0, 0), datetime.datetime(2026, 1, 3, 0, 0), datetime.datetime(2026, 2, 4, 0, 0)]
    <- Calc!N2:N61 (expanded above)
  <- Calc!P1:P61 =IF(ROUNDUP(MONTH($E2)/3,0)=2,$N2,0)  (x60)
        ! 1 typed value(s) inside this formula run, rows 1  <~ dynamic
    <- Calc!E2:E61 (expanded above)
    <- Calc!N2:N61 (expanded above)
  <- Calc!Q1:Q61 =IF(ROUNDUP(MONTH($E2)/3,0)=3,$N2,0)  (x60)
        ! 1 typed value(s) inside this formula run, rows 1  <~ dynamic
    <- Calc!E2:E61 (expanded above)
    <- Calc!N2:N61 (expanded above)
  <- Calc!R1:R61 =IF(ROUNDUP(MONTH($E2)/3,0)=4,$N2,0)  (x60)
        ! 1 typed value(s) inside this formula run, rows 1  <~ dynamic
    <- Calc!E2:E61 (expanded above)
    <- Calc!N2:N61 (expanded above)
  <- Summary!A5:A16 (expanded above)
  <- Summary!B2 typed values (1) e.g. ['Q3']
```

### Summary!E5:E16

```
Summary!E5:E16 =COUNTIF(Calc!$B$2:$B$61,$A5)  (x12)
  <- Calc!B2:B61 (expanded above)
  <- Summary!A5:A16 (expanded above)
```

## True inputs

| Block | Header | Cells | Types | Range / values | Thresholds downstream |
|---|---|---|---|---|---|
| Calc!O1 |  | 1 | {'text': 1} | 'Q1' |  |
| Calc!P1 |  | 1 | {'text': 1} | 'Q2' |  |
| Calc!Q1 |  | 1 | {'text': 1} | 'Q3' |  |
| Calc!R1 |  | 1 | {'text': 1} | 'Q4' |  |
| Calc!S40:S61 | Final Commission | 22 | {'number': 22} | 0 .. 205,404.37499999997 |  |
| Inputs!B2:B61 | Sales Rep | 60 | {'text': 60} | 'Amelie Laurent', 'Ananya Iyer', 'Daniel Kim', 'Hiroshi Tanaka', 'Jessica Alvarez', 'Lukas Becker' ... | 1, 2, 3, 4, 90 |
| Inputs!C2:C61 | Region | 60 | {'text': 60} | 'AMER', 'APAC', 'EMEA', 'India' | 25000, 50000, 100000, 150000, 250000, 400000, 600000, 1e+06 |
| Inputs!D2:D61 | Product | 60 | {'text': 60} | 'Analytics Add-on', 'Core Platform', 'Data Connectors', 'Enterprise Suite', 'Support Plus' |  |
| Inputs!E2:E61 | Deal Value | 60 | {'number': 56, 'text': 4} | 23,500 .. 1,082,500 | 1, 2, 3, 4, 90, 25000, 50000, 100000 |
| Inputs!G2:G61 | Close Date | 60 | {'date': 60} | 2026-01-01T00:00:00 .. 2026-12-28T00:00:00 | 1, 2, 3, 4 |
| Inputs!H2:H61 | Days to Pay | 60 | {'number': 60} | 20 .. 130 | 1, 2, 3, 4, 90 |
| Rates!A2:A10 | Deal Size From (USD) | 9 | {'number': 9} | 0 .. 1,000,000 | 90 |
| Rates!B2:B10 | Base Rate | 9 | {'number': 9} | 0.02 .. 0.11 | 90 |
| Rates!E2:E10 | Tier Kicker | 9 | {'number': 9} | 0 .. 0.025 | 90 |
| Rates!G2:G6 | Product | 5 | {'text': 5} | 'Analytics Add-on', 'Core Platform', 'Data Connectors', 'Enterprise Suite', 'Support Plus' | 90 |
| Rates!H2:H6 | Product Bonus | 5 | {'number': 5} | 0 .. 0.04 | 90 |
| Rates!J2:J13 | Sales Rep | 12 | {'text': 12} | 'Amelie Laurent', 'Ananya Iyer', 'Daniel Kim', 'Hiroshi Tanaka', 'Jessica Alvarez', 'Lukas Becker' ... | 1, 2, 3, 4, 90 |
| Rates!K2:K13 | Plan Participation | 12 | {'number': 12} | 0.5 .. 1 | 1, 2, 3, 4, 90 |
| Summary!A5:A16 | Sales Rep | 12 | {'text': 12} | 'Amelie Laurent', 'Ananya Iyer', 'Daniel Kim', 'Hiroshi Tanaka', 'Jessica Alvarez', 'Lukas Becker' ... |  |
| Summary!A19:A22 | Region | 4 | {'text': 4} | 'AMER', 'APAC', 'EMEA', 'India' |  |
| Summary!B2 |  | 1 | {'text': 1} | 'Q3' |  |
| adj!A2:A3 (veryHidden) | Rep | 2 | {'text': 2} | "Sarah O'Connor", 'Wei Zhang' |  |
| adj!B2:B3 (veryHidden) | Amount | 2 | {'number': 2} | -1,200 .. 4,500 |  |

## Dynamic references (12)

- cell: Summary!D5; function: OFFSET; formula: =SUMIF(Calc!$B$2:$B$61,$A5,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1)); base: ['Calc!O2:O61']; drivers: ['Summary!B2', 'Calc!O1:R1']; candidates: ['Calc!O2:R61']
- cell: Summary!D6; function: OFFSET; formula: =SUMIF(Calc!$B$2:$B$61,$A6,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1)); base: ['Calc!O2:O61']; drivers: ['Summary!B2', 'Calc!O1:R1']; candidates: ['Calc!O2:R61']
- cell: Summary!D7; function: OFFSET; formula: =SUMIF(Calc!$B$2:$B$61,$A7,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1)); base: ['Calc!O2:O61']; drivers: ['Summary!B2', 'Calc!O1:R1']; candidates: ['Calc!O2:R61']
- cell: Summary!D8; function: OFFSET; formula: =SUMIF(Calc!$B$2:$B$61,$A8,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1)); base: ['Calc!O2:O61']; drivers: ['Summary!B2', 'Calc!O1:R1']; candidates: ['Calc!O2:R61']
- cell: Summary!D9; function: OFFSET; formula: =SUMIF(Calc!$B$2:$B$61,$A9,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1)); base: ['Calc!O2:O61']; drivers: ['Summary!B2', 'Calc!O1:R1']; candidates: ['Calc!O2:R61']
- cell: Summary!D10; function: OFFSET; formula: =SUMIF(Calc!$B$2:$B$61,$A10,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1)); base: ['Calc!O2:O61']; drivers: ['Summary!B2', 'Calc!O1:R1']; candidates: ['Calc!O2:R61']
- cell: Summary!D11; function: OFFSET; formula: =SUMIF(Calc!$B$2:$B$61,$A11,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1)); base: ['Calc!O2:O61']; drivers: ['Summary!B2', 'Calc!O1:R1']; candidates: ['Calc!O2:R61']
- cell: Summary!D12; function: OFFSET; formula: =SUMIF(Calc!$B$2:$B$61,$A12,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1)); base: ['Calc!O2:O61']; drivers: ['Summary!B2', 'Calc!O1:R1']; candidates: ['Calc!O2:R61']
- cell: Summary!D13; function: OFFSET; formula: =SUMIF(Calc!$B$2:$B$61,$A13,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1)); base: ['Calc!O2:O61']; drivers: ['Summary!B2', 'Calc!O1:R1']; candidates: ['Calc!O2:R61']
- cell: Summary!D14; function: OFFSET; formula: =SUMIF(Calc!$B$2:$B$61,$A14,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1)); base: ['Calc!O2:O61']; drivers: ['Summary!B2', 'Calc!O1:R1']; candidates: ['Calc!O2:R61']
- cell: Summary!D15; function: OFFSET; formula: =SUMIF(Calc!$B$2:$B$61,$A15,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1)); base: ['Calc!O2:O61']; drivers: ['Summary!B2', 'Calc!O1:R1']; candidates: ['Calc!O2:R61']
- cell: Summary!D16; function: OFFSET; formula: =SUMIF(Calc!$B$2:$B$61,$A16,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1)); base: ['Calc!O2:O61']; drivers: ['Summary!B2', 'Calc!O1:R1']; candidates: ['Calc!O2:R61']

## Broken references (0)


## Lookup tables with problems (1)

- function: VLOOKUP; table: Rates!A2:B10; approximate: True; problems: ['approximate match on keys that are NOT sorted ascending (first break after position 4: 150000 > 100000)']; count: 60; cells: ['Calc!H2', 'Calc!H3', 'Calc!H4', 'Calc!H5', 'Calc!H6', 'Calc!H7', 'Calc!H8', 'Calc!H9']

## Aggregates that stop short of adjacent data (0)


## Circular references (0)


## Other end-points (60 formula cells nothing reads, not on an output sheet)

Calc!A2, Calc!A3, Calc!A4, Calc!A5, Calc!A6, Calc!A7, Calc!A8, Calc!A9, Calc!A10, Calc!A11, Calc!A12, Calc!A13, Calc!A14, Calc!A15, Calc!A16, Calc!A17, Calc!A18, Calc!A19, Calc!A20, Calc!A21, Calc!A22, Calc!A23, Calc!A24, Calc!A25, Calc!A26, Calc!A27, Calc!A28, Calc!A29, Calc!A30, Calc!A31, Calc!A32, Calc!A33, Calc!A34, Calc!A35, Calc!A36, Calc!A37, Calc!A38, Calc!A39, Calc!A40, Calc!A41, Calc!A42, Calc!A43, Calc!A44, Calc!A45, Calc!A46, Calc!A47, Calc!A48, Calc!A49, Calc!A50, Calc!A51, Calc!A52, Calc!A53, Calc!A54, Calc!A55, Calc!A56, Calc!A57, Calc!A58, Calc!A59, Calc!A60, Calc!A61
