#!/usr/bin/env python3
# Test the updated Array Number logic

import pandas as pd

# Test case từ ví dụ của bạn
cross_passage = "EXP61103"
array_number = "EXP61103B"

print("Testing updated Array Number logic:")
print(f"Cross Passage: '{cross_passage}'")
print(f"Array Number: '{array_number}'")
print()

# New logic: Array Number should contain Cross Passage
cross_passage_str = str(cross_passage).strip()
actual_array = str(array_number).strip()

result = cross_passage_str in actual_array
print(f"Logic: '{cross_passage_str}' in '{actual_array}' = {result}")

if result:
    print("✅ PASS - Array Number chứa Cross Passage!")
else:
    print(f"❌ FAIL - cần chứa '{cross_passage_str}', có '{actual_array}'")

print()
print("Test các trường hợp khác:")

test_cases = [
    {"cross": "EXP61003", "array": "EXP61003", "should_pass": True},
    {"cross": "EXP61003", "array": "EXP61003X", "should_pass": True},  
    {"cross": "EXP61103", "array": "EXP61103B", "should_pass": True},
    {"cross": "EXP61103", "array": "EXP61203", "should_pass": False},
]

for i, case in enumerate(test_cases, 1):
    cross = case["cross"]
    array = case["array"]
    expected = case["should_pass"]
    
    result = cross in array
    status = "✅ PASS" if result == expected else "❌ FAIL"
    
    print(f"Test {i}: '{cross}' in '{array}' = {result} {status}")
