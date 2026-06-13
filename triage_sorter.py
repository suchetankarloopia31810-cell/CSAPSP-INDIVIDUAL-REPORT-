"""
Hospital Emergency Department (ED) Patient Triage Management System
Module : Algorithms and Problem Solving using Python
Assignment: Transforming a Real-World Problem into an Algorithm using Python

Scenario:
  An NHS emergency department receives patient arrivals around the clock.
  Each patient is assigned a triage score (1–5) based on the Manchester
  Triage System, where 1 = Immediate (life-threatening) and 5 = Non-Urgent.
  The system must sort the patient queue in ascending order of triage score
  so that the most critical patients are seen first, flag all Level-1 and
  Level-2 patients for immediate clinical attention, and report performance
  metrics to assist operational benchmarking.

  The system must scale from a small shift log (~7–20 patients) to a full
  monthly export (~10,000 patient records).
"""

import time
import random
from copy import deepcopy

# ─────────────────────────────────────────────────────────────────────
# Manchester Triage System labels
# ─────────────────────────────────────────────────────────────────────
TRIAGE_LABEL = {
    1: "Immediate  (Red)",
    2: "Very Urgent (Orange)",
    3: "Urgent     (Yellow)",
    4: "Standard   (Green)",
    5: "Non-Urgent (Blue)",
}


# ─────────────────────────────────────────────────────────────────────
# 1.  Operation Counter Class
# ─────────────────────────────────────────────────────────────────────
class OpCounter:
    """
    Tracks every arithmetic step, element comparison and swap performed
    during a sort so that empirical operation counts can be compared
    against theoretical Big-O predictions.
    """

    def __init__(self):
        self.comparisons = 0
        self.swaps = 0
        self.arithmetic = 0

    def total(self):
        return self.comparisons + self.swaps + self.arithmetic

    def reset(self):
        self.comparisons = 0
        self.swaps = 0
        self.arithmetic = 0


# ─────────────────────────────────────────────────────────────────────
# 2.  Sorting Algorithms
# ─────────────────────────────────────────────────────────────────────

def bubble_sort(records, counter):
    """
    Bubble Sort  –  O(n²) time / O(1) auxiliary space.
    Repeatedly scans the list, swapping adjacent patient records that are
    out of triage priority order. Each full pass moves the least-urgent
    patient to the end of the unsorted portion.
    """
    data = deepcopy(records)
    n = len(data)
    for i in range(n):
        counter.arithmetic += 1                         # outer loop increment
        for j in range(0, n - i - 1):
            counter.arithmetic += 1                     # inner loop increment
            counter.comparisons += 1                    # triage score comparison
            if data[j]["score"] > data[j + 1]["score"]:
                data[j], data[j + 1] = data[j + 1], data[j]
                counter.swaps += 1
    return data


def insertion_sort(records, counter):
    """
    Insertion Sort  –  O(n²) time / O(1) auxiliary space.
    Builds the sorted queue one patient at a time by inserting each new
    arrival into the correct position among already-prioritised patients.
    Performs well on nearly-sorted (shift-handover) data.
    """
    data = deepcopy(records)
    for i in range(1, len(data)):
        counter.arithmetic += 1                         # loop increment
        key = data[i]
        j = i - 1
        while j >= 0:
            counter.comparisons += 1
            counter.arithmetic += 1                     # j decrement
            if data[j]["score"] > key["score"]:
                data[j + 1] = data[j]
                counter.swaps += 1
                j -= 1
            else:
                break
        data[j + 1] = key
    return data


def _merge_helper(data, counter):
    """Recursive divide-and-merge helper for Merge Sort."""
    if len(data) <= 1:
        return
    mid = len(data) // 2
    counter.arithmetic += 1                             # midpoint calculation
    left = data[:mid]
    right = data[mid:]
    _merge_helper(left, counter)
    _merge_helper(right, counter)
    i = j = k = 0
    while i < len(left) and j < len(right):
        counter.comparisons += 1
        counter.arithmetic += 1
        if left[i]["score"] <= right[j]["score"]:
            data[k] = left[i]
            i += 1
        else:
            data[k] = right[j]
            j += 1
        k += 1
        counter.arithmetic += 1
    while i < len(left):
        data[k] = left[i]
        i += 1
        k += 1
        counter.arithmetic += 1
    while j < len(right):
        data[k] = right[j]
        j += 1
        k += 1
        counter.arithmetic += 1


def merge_sort(records, counter):
    """
    Merge Sort  –  O(n log n) time / O(n) auxiliary space.
    Recursively divides the patient queue, sorts each half independently,
    then merges them in score order. Guarantees O(n log n) in all cases —
    critical when patient volumes are unpredictable.
    """
    data = deepcopy(records)
    _merge_helper(data, counter)
    return data


def _three_way_partition(data, low, high, counter):
    """
    Dutch National Flag (3-way) partition around a randomly chosen pivot.
    Returns (lt, gt) such that:
      data[low..lt-1]  < pivot
      data[lt..gt]    == pivot
      data[gt+1..high] > pivot
    Randomised pivot selection avoids worst-case O(n²) on repeated values,
    which is critical for triage data where only 5 distinct scores exist.
    """
    # Randomly select pivot and move to low position
    rand_idx = random.randint(low, high)
    data[low], data[rand_idx] = data[rand_idx], data[low]
    counter.swaps += 1

    pivot = data[low]["score"]
    lt = low       # data[low..lt-1]  < pivot
    gt = high      # data[gt+1..high] > pivot
    i  = low + 1   # current element under inspection

    counter.arithmetic += 3  # lt, gt, i initialisations

    while i <= gt:
        counter.comparisons += 1
        counter.arithmetic  += 1
        if data[i]["score"] < pivot:
            data[lt], data[i] = data[i], data[lt]
            counter.swaps += 1
            lt += 1
            i  += 1
            counter.arithmetic += 2
        elif data[i]["score"] > pivot:
            data[i], data[gt] = data[gt], data[i]
            counter.swaps += 1
            gt -= 1
            counter.arithmetic += 1
        else:
            i += 1
            counter.arithmetic += 1

    return lt, gt


def _quick_helper(data, low, high, counter):
    if low < high:
        counter.comparisons += 1
        lt, gt = _three_way_partition(data, low, high, counter)
        _quick_helper(data, low,    lt - 1, counter)
        _quick_helper(data, gt + 1, high,   counter)


def quick_sort(records, counter):
    """
    Quick Sort (3-way / randomised pivot)  –  O(n log n) expected time / O(log n) space.
    Uses a Dutch National Flag partition with a randomly chosen pivot, giving
    O(n log n) expected performance even when triage scores are highly repeated
    (only 5 distinct values in 10,000 records).
    """
    data = deepcopy(records)
    _quick_helper(data, 0, len(data) - 1, counter)
    return data


# ─────────────────────────────────────────────────────────────────────
# 3.  Helper Utilities
# ─────────────────────────────────────────────────────────────────────

def generate_patients(size, seed=None):
    """Generates a synthetic ED patient dataset with random triage scores."""
    if seed is not None:
        random.seed(seed)
    first_names = [
        "Aisha", "Ben", "Clara", "Dion", "Elena", "Farid", "Grace",
        "Haruto", "Ingrid", "Jorge", "Kofi", "Leila", "Marcus", "Nadia",
        "Owen", "Petra", "Quentin", "Rosa", "Stefan", "Tiana"
    ]
    patients = []
    for idx in range(size):
        name = random.choice(first_names) + f"_{idx:05d}"
        score = random.randint(1, 5)
        patients.append({"name": name, "score": score})
    return patients


def flag_critical(sorted_records, threshold=2):
    """Returns all patients whose triage score is at or below the threshold."""
    return [p for p in sorted_records if p["score"] <= threshold]


def run_sort(func, records):
    """Runs a sort function, times it and returns (sorted_data, ms, counter)."""
    ctr = OpCounter()
    t0 = time.perf_counter()
    result = func(deepcopy(records), ctr)
    elapsed = (time.perf_counter() - t0) * 1000
    return result, elapsed, ctr


# ─────────────────────────────────────────────────────────────────────
# 4.  Main Execution
# ─────────────────────────────────────────────────────────────────────
ALGORITHMS = [
    ("Bubble Sort",    bubble_sort),
    ("Insertion Sort", insertion_sort),
    ("Merge Sort",     merge_sort),
    ("Quick Sort",     quick_sort),
]

HDR = (f"{'Algorithm':<20} {'Time (ms)':>11} {'Comparisons':>13} "
       f"{'Swaps':>8} {'Arithmetic':>12} {'Total Ops':>11}")
SEP = "─" * 79

if __name__ == "__main__":

    # ── Fixed small dataset (assignment scenario) ──────────────────────
    small = [
        {"name": "James",   "score": 3},
        {"name": "Maria",   "score": 1},
        {"name": "Chen",    "score": 4},
        {"name": "Fatima",  "score": 2},
        {"name": "Oliver",  "score": 5},
        {"name": "Priya",   "score": 1},
        {"name": "Marcus",  "score": 3},
    ]

    print("=" * 60)
    print("HOSPITAL ED PATIENT TRIAGE MANAGEMENT SYSTEM")
    print("=" * 60)

    # Small
    print(f"\n[SMALL DATASET]  {len(small)} patients")
    print(HDR); print(SEP)
    ms_sorted_small = None
    for name, fn in ALGORITHMS:
        sd, ms, ctr = run_sort(fn, small)
        print(f"{name:<20} {ms:>11.4f} {ctr.comparisons:>13} "
              f"{ctr.swaps:>8} {ctr.arithmetic:>12} {ctr.total():>11}")
        if name == "Merge Sort":
            ms_sorted_small = sd

    print("\nSorted queue (Merge Sort – ascending triage score):")
    for p in ms_sorted_small:
        print(f"  {p['name']:<10}  Score {p['score']}  –  {TRIAGE_LABEL[p['score']]}")

    critical_small = flag_critical(ms_sorted_small)
    print(f"\nCritical patients flagged (score ≤ 2): {len(critical_small)}")
    for p in critical_small:
        print(f"  ** {p['name']:<10}  Score {p['score']}  –  {TRIAGE_LABEL[p['score']]}")

    # Medium
    medium = generate_patients(50, seed=7)
    print(f"\n[MEDIUM DATASET]  {len(medium)} patients")
    print(HDR); print(SEP)
    ms_sorted_medium = None
    for name, fn in ALGORITHMS:
        sd, ms, ctr = run_sort(fn, medium)
        print(f"{name:<20} {ms:>11.4f} {ctr.comparisons:>13} "
              f"{ctr.swaps:>8} {ctr.arithmetic:>12} {ctr.total():>11}")
        if name == "Merge Sort":
            ms_sorted_medium = sd
    print(f"Critical patients: {len(flag_critical(ms_sorted_medium))} / {len(medium)}")

    # Large
    large = generate_patients(10000, seed=42)
    print(f"\n[LARGE DATASET]  {len(large)} patients")
    print(HDR); print(SEP)
    ms_sorted_large = None
    for name, fn in ALGORITHMS:
        sd, ms, ctr = run_sort(fn, large)
        print(f"{name:<20} {ms:>11.4f} {ctr.comparisons:>13} "
              f"{ctr.swaps:>8} {ctr.arithmetic:>12} {ctr.total():>11}")
        if name == "Merge Sort":
            ms_sorted_large = sd
    print(f"Critical patients: {len(flag_critical(ms_sorted_large))} / {len(large)}")
    print("\nDone.")
