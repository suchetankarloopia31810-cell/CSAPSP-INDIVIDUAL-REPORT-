"""
Library Management System – Borrowing Records Sorter
Module : Algorithms and Problem Solving using Python
Assignment: Transforming a Real-World Problem into an Algorithm using Python

This program reads/generates library borrowing records, sorts them by days
using four algorithms (Bubble, Insertion, Merge, Quick Sort), counts every
arithmetic, comparison and swap operation, measures execution time, and
identifies overdue borrowers (> 14 days).
"""

import time
import random
import string
from copy import deepcopy


# ===========================================================
# 1.  Operation Counter Class
# ===========================================================
class OpCounter:
    """Tracks comparisons, swaps, and arithmetic operations performed
    during a sort so that algorithm complexity can be verified empirically."""

    def __init__(self):
        self.comparisons = 0
        self.swaps       = 0
        self.arithmetic  = 0

    def total(self):
        return self.comparisons + self.swaps + self.arithmetic

    def reset(self):
        self.comparisons = 0
        self.swaps       = 0
        self.arithmetic  = 0


# ===========================================================
# 2.  Sorting Algorithms
# ===========================================================

def bubble_sort(records, counter):
    """
    Bubble Sort  –  O(n²) time / O(1) space.
    Repeatedly scans the list, swapping adjacent elements that are
    out of order. Each full pass 'bubbles' the largest unsorted value
    to the end of the list.
    """
    data = deepcopy(records)
    n = len(data)
    for i in range(n):
        counter.arithmetic += 1                         # loop variable update
        for j in range(0, n - i - 1):
            counter.arithmetic += 1                     # inner loop variable update
            counter.comparisons += 1                    # compare days values
            if data[j]['days'] > data[j + 1]['days']:
                data[j], data[j + 1] = data[j + 1], data[j]
                counter.swaps += 1
    return data


def insertion_sort(records, counter):
    """
    Insertion Sort  –  O(n²) time / O(1) space.
    Builds the sorted portion one record at a time by inserting each
    unsorted record into its correct position among already-sorted records.
    """
    data = deepcopy(records)
    for i in range(1, len(data)):
        counter.arithmetic += 1                         # loop variable update
        key = data[i]
        j = i - 1
        while j >= 0:
            counter.comparisons += 1
            counter.arithmetic  += 1                    # j decrement
            if data[j]['days'] > key['days']:
                data[j + 1] = data[j]
                counter.swaps += 1
                j -= 1
            else:
                break
        data[j + 1] = key
    return data


def _merge_sort_helper(data, counter):
    """Recursive helper that splits, sorts and merges sub-arrays in place."""
    if len(data) <= 1:
        return
    mid   = len(data) // 2
    counter.arithmetic += 1                             # mid calculation
    left  = data[:mid]
    right = data[mid:]

    _merge_sort_helper(left,  counter)
    _merge_sort_helper(right, counter)

    i = j = k = 0
    while i < len(left) and j < len(right):
        counter.comparisons += 1
        counter.arithmetic  += 1
        if left[i]['days'] <= right[j]['days']:
            data[k] = left[i]
            i += 1
        else:
            data[k] = right[j]
            j += 1
        k += 1
        counter.arithmetic += 1                         # k increment
    while i < len(left):
        data[k] = left[i]
        i += 1; k += 1
        counter.arithmetic += 1
    while j < len(right):
        data[k] = right[j]
        j += 1; k += 1
        counter.arithmetic += 1


def merge_sort(records, counter):
    """
    Merge Sort  –  O(n log n) time / O(n) space.
    Divides the list into halves recursively, sorts each half, then merges
    them back in sorted order. Guarantees O(n log n) in all cases.
    """
    data = deepcopy(records)
    _merge_sort_helper(data, counter)
    return data


def _partition(data, low, high, counter):
    """Lomuto partition scheme – pivot is the last element."""
    pivot = data[high]['days']
    i = low - 1
    counter.arithmetic += 1                             # i initialisation
    for j in range(low, high):
        counter.comparisons += 1
        counter.arithmetic  += 1                        # j increment
        if data[j]['days'] <= pivot:
            i += 1
            data[i], data[j] = data[j], data[i]
            counter.swaps    += 1
            counter.arithmetic += 1                     # i increment
    data[i + 1], data[high] = data[high], data[i + 1]
    counter.swaps     += 1
    counter.arithmetic += 1
    return i + 1


def _quick_sort_helper(data, low, high, counter):
    if low < high:
        counter.comparisons += 1
        pi = _partition(data, low, high, counter)
        _quick_sort_helper(data, low,    pi - 1, counter)
        _quick_sort_helper(data, pi + 1, high,   counter)


def quick_sort(records, counter):
    """
    Quick Sort  –  O(n log n) average / O(n²) worst-case time / O(log n) space.
    Selects a pivot element and partitions the list so that smaller values
    precede the pivot and larger values follow it, then recursively sorts each
    partition.
    """
    data = deepcopy(records)
    _quick_sort_helper(data, 0, len(data) - 1, counter)
    return data


# ===========================================================
# 3.  Utility Functions
# ===========================================================

def generate_dataset(size):
    """Generates a synthetic dataset of random borrower records."""
    first_names = [
        'Alice', 'Bob', 'Carol', 'David', 'Eve', 'Frank', 'Grace',
        'Henry', 'Isla', 'James', 'Karen', 'Leo', 'Mia', 'Nora',
        'Oscar', 'Pam', 'Quinn', 'Rex', 'Sara', 'Tom'
    ]
    records = []
    for _ in range(size):
        name = random.choice(first_names) + '_' + ''.join(
            random.choices(string.digits, k=4))
        records.append({'name': name, 'days': random.randint(1, 60)})
    return records


def find_late_returns(sorted_records, overdue_limit=14):
    """Returns all records where days exceed the overdue limit."""
    return [r for r in sorted_records if r['days'] > overdue_limit]


def run_algorithm(name, func, records):
    """Executes a sorting function, times it and returns performance metrics."""
    counter = OpCounter()
    start       = time.perf_counter()
    sorted_data = func(deepcopy(records), counter)
    elapsed_ms  = (time.perf_counter() - start) * 1000
    return sorted_data, elapsed_ms, counter


# ===========================================================
# 4.  Main Execution
# ===========================================================
if __name__ == '__main__':

    # ── Fixed small dataset from the assignment brief ──────────────────────
    small_dataset = [
        {'name': 'Emma',   'days': 23},
        {'name': 'Liam',   'days':  5},
        {'name': 'Sophia', 'days': 31},
        {'name': 'Mason',  'days': 14},
        {'name': 'Olivia', 'days':  7},
        {'name': 'Noah',   'days': 42},
        {'name': 'Ava',    'days': 18},
    ]

    ALGORITHMS = [
        ('Bubble Sort',    bubble_sort),
        ('Insertion Sort', insertion_sort),
        ('Merge Sort',     merge_sort),
        ('Quick Sort',     quick_sort),
    ]

    HDR = f"{'Algorithm':<20} {'Time (ms)':>12} {'Comparisons':>14} " \
          f"{'Swaps':>8} {'Arithmetic':>12} {'Total Ops':>12}"
    SEP = '-' * 80

    # ── Small dataset ───────────────────────────────────────────────────────
    print('=' * 60)
    print('LIBRARY MANAGEMENT SYSTEM – BORROWING RECORDS SORTER')
    print('=' * 60)
    print(f'\n[SMALL DATASET]  {len(small_dataset)} records')
    print(HDR); print(SEP)

    merge_sorted_small = None
    for alg_name, alg_func in ALGORITHMS:
        sd, ms, ctr = run_algorithm(alg_name, alg_func, small_dataset)
        print(f"{alg_name:<20} {ms:>12.4f} {ctr.comparisons:>14} "
              f"{ctr.swaps:>8} {ctr.arithmetic:>12} {ctr.total():>12}")
        if alg_name == 'Merge Sort':
            merge_sorted_small = sd

    print('\nSorted Order (Merge Sort – ascending days):')
    for r in merge_sorted_small:
        print(f"  {r['name']:<10}  {r['days']:>3} days")

    late_small = find_late_returns(merge_sorted_small)
    print(f'\nOverdue borrowers (> 14 days): {len(late_small)}')
    for r in late_small:
        print(f"  {r['name']:<10}  {r['days']:>3} days")

    # ── Medium dataset ──────────────────────────────────────────────────────
    random.seed(42)
    medium_dataset = generate_dataset(50)
    print(f'\n[MEDIUM DATASET]  {len(medium_dataset)} records')
    print(HDR); print(SEP)

    merge_sorted_medium = None
    for alg_name, alg_func in ALGORITHMS:
        sd, ms, ctr = run_algorithm(alg_name, alg_func, medium_dataset)
        print(f"{alg_name:<20} {ms:>12.4f} {ctr.comparisons:>14} "
              f"{ctr.swaps:>8} {ctr.arithmetic:>12} {ctr.total():>12}")
        if alg_name == 'Merge Sort':
            merge_sorted_medium = sd

    late_medium = find_late_returns(merge_sorted_medium)
    print(f'Overdue borrowers: {len(late_medium)} / {len(medium_dataset)}')

    # ── Large dataset ────────────────────────────────────────────────────────
    random.seed(99)
    large_dataset = generate_dataset(10000)
    print(f'\n[LARGE DATASET]  {len(large_dataset)} records')
    print(HDR); print(SEP)

    merge_sorted_large = None
    for alg_name, alg_func in ALGORITHMS:
        sd, ms, ctr = run_algorithm(alg_name, alg_func, large_dataset)
        print(f"{alg_name:<20} {ms:>12.4f} {ctr.comparisons:>14} "
              f"{ctr.swaps:>8} {ctr.arithmetic:>12} {ctr.total():>12}")
        if alg_name == 'Merge Sort':
            merge_sorted_large = sd

    late_large = find_late_returns(merge_sorted_large)
    print(f'Overdue borrowers: {len(late_large)} / {len(large_dataset)}')
    print('\nDone.')
