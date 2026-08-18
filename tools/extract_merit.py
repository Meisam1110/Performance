#!/usr/bin/env python3
"""Extract sample data + Excel expected-value fixtures from Merit.xlsb.

Produces:
  sample-data/merit-sample.json   -> input data (questionnaire + master) for the web app
  tests/merit-expected.json       -> every computed cell Excel produced, for regression tests
"""
import json, sys, os
from pyxlsb import open_workbook

SRC = sys.argv[1] if len(sys.argv) > 1 else 'Merit.xlsb'
OUT_SAMPLE = 'sample-data/merit-sample.json'
OUT_EXPECT = 'tests/merit-expected.json'

FIRST, LAST = 8, 107          # data rows on both calculation sheets (1-based)
Q_SHEET   = 'پرسشنامه کارانه تیمی'
PAY_SHEET = 'روش پرداخت کارانه'
DATA_SHEET = 'Data'


def read(name, wb):
    with wb.get_sheet(name) as sh:
        return {(c.r + 1, c.c): c.v for row in sh.rows() for c in row if c.v not in (None, '')}


def jl(v):
    """Job level comes through as float 3.0 or string '3H'."""
    if isinstance(v, str):
        return v.strip()
    return str(int(v))


def main():
    wb = open_workbook(SRC)
    Q, P, D = read(Q_SHEET, wb), read(PAY_SHEET, wb), read(DATA_SHEET, wb)

    # --- configuration lifted out of the workbook itself -------------------
    grade_map = {}
    for r in range(4, 10):                       # Data!C4:D9
        k, v = D.get((r, 2)), D.get((r, 3))
        if k is not None and v is not None:
            grade_map[jl(k)] = v
    answer_scale = {}
    for r in range(4, 9):                        # Data!H4:I8
        k, v = D.get((r, 7)), D.get((r, 8))
        if k is not None and v is not None:
            answer_scale[k] = v
    divisions = [D[(r, 1)] for r in range(1, 9) if (r, 1) in D]

    config = {
        'budget': P[(1, 2)],                     # روش پرداخت کارانه!C1
        'gradeImpactFactor': P.get((2, 3), 0),   # D2
        'minPerformanceThreshold': P[(4, 3)],    # D4
        'gradeMap': grade_map,
        'answerScale': answer_scale,
        'maxPerformanceScore': 120,
        'questionCount': 4,
        'baselineCoefficientPerPerson': Q[(5, 1)] / Q[(5, 0)],   # B5 / A5
        'specialImpactAmount': 300,
        'divisions': divisions,
    }

    questions = [D[(r, 0)] for r in range(3, 9) if (r, 0) in read('Q', wb)] if False else None
    qtexts = read('Q', wb)
    config['questionTexts'] = [qtexts[(r, 0)] for r in range(3, 9) if (r, 0) in qtexts]
    config['questionHeaders'] = [Q.get((6, c)) for c in range(5, 10)] + [Q.get((6, 12))]

    employees, expected = [], []
    for r in range(FIRST, LAST + 1):
        if (r, 0) not in Q:
            continue
        employees.append({
            'employeeId':   str(int(Q[(r, 0)])) if isinstance(Q[(r, 0)], float) else str(Q[(r, 0)]),
            'fullName':     Q.get((r, 1)),
            'division':     Q.get((r, 2)),
            'positionTitle': Q.get((r, 3)),
            'jobLevel':     jl(Q.get((r, 4))),
            'q1': Q.get((r, 5)), 'q2': Q.get((r, 6)), 'q3': Q.get((r, 7)),
            'q4': Q.get((r, 8)), 'q5': Q.get((r, 9)),
            'specialProject': Q.get((r, 12)) or None,
            'specialImpactAmount': Q.get((r, 13)) or 0,
            'hodAdjustment': P.get((r, 13)),
            'sourceFile': 'Merit.xlsb',
        })
        expected.append({
            'employeeId': employees[-1]['employeeId'],
            # --- پرسشنامه کارانه تیمی ---
            'performanceScore':        Q.get((r, 10)),   # K
            'performanceKaraneh':      Q.get((r, 11)),   # L
            'specialImpactValue':      Q.get((r, 13), 0) or 0,  # N
            'rawCoefficient':          Q.get((r, 14)),   # O
            'finalCoefficient':        Q.get((r, 17)),   # R
            # --- روش پرداخت کارانه ---
            'gradeScore':              P.get((r, 5)),    # F
            'gradeImpact':             P.get((r, 6), 0), # G
            'evalScore':               P.get((r, 7), 0), # H
            'eligibleEvalScore':       P.get((r, 8), 0), # I
            'baseCoefficient':         P.get((r, 9), 0), # J
            'performanceScoreFinal':   P.get((r, 10), 0),# K
            'totalScore':              P.get((r, 11), 0),# L
            'initialAllocation':       P.get((r, 12), 0),# M
            'hodAdjustment':           P.get((r, 13)),   # N
            'diff':                    P.get((r, 14)),   # O
            'finalKaraneh':            P.get((r, 16), 0),# Q
        })

    totals = {
        'questionnaireCount':        Q[(5, 0)],
        'baselineTotal':             Q[(5, 1)],
        'coefficientExcess':         Q[(5, 14)],
        'perPersonNormalization':    Q[(5, 13)],
        'ineligibleRedistribution':  P[(5, 9)],
        'hodRedistribution':         P[(5, 14)],
        'sumInitialAllocation':      P[(5, 12)],
        'sumFinalKaraneh':           P[(5, 16)],
    }

    os.makedirs('sample-data', exist_ok=True)
    os.makedirs('tests', exist_ok=True)
    with open(OUT_SAMPLE, 'w', encoding='utf-8') as f:
        json.dump({'config': config, 'employees': employees}, f, ensure_ascii=False, indent=1)
    with open(OUT_EXPECT, 'w', encoding='utf-8') as f:
        json.dump({'totals': totals, 'rows': expected}, f, ensure_ascii=False, indent=1)
    print('wrote %s (%d employees) and %s' % (OUT_SAMPLE, len(employees), OUT_EXPECT))


if __name__ == '__main__':
    main()
