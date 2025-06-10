#!/usr/bin/env python3
# Test logic Array Number validation

# Test case từ ví dụ của bạn
cross_passage = 'B'
location_lanes = '03' 
array_number = 'EXP61103B'

# Theo logic enhanced validator
import re

cross_passage_str = str(cross_passage)
location_lanes_str = str(location_lanes)

# Extract 2 digits cuối
cross_digits = re.findall(r'\d', cross_passage_str)[-2:] if len(re.findall(r'\d', cross_passage_str)) >= 2 else ['0', '0']
location_digits = re.findall(r'\d', location_lanes_str)[-2:] if len(re.findall(r'\d', location_lanes_str)) >= 2 else ['0', '0']

# Tạo expected array number
expected_array = f"EXP6{''.join(location_digits)}{''.join(cross_digits)}"
actual_array = str(array_number).strip()

print(f"Cross Passage: {cross_passage}")
print(f"Location Lanes: {location_lanes}")
print(f"Array Number: {array_number}")
print(f"Expected Pattern: {expected_array}")
print(f"Old Logic (==): {actual_array == expected_array}")
print(f"New Logic (in): {expected_array in actual_array}")
print()

if expected_array in actual_array:
    print("✅ PASS - Array Number chứa pattern mong đợi!")
else:
    print("❌ FAIL - Array Number không chứa pattern mong đợi!")
