#!/usr/bin/env python3
# Test to understand the correct digit extraction logic

# From your example:
# Cross Passage: probably contains "03" (last 2 digits)
# Location Lanes: probably contains "11" (last 2 digits)  
# Expected: EXP61103
# Actual: EXP61103B

# Let's reverse engineer:
expected = "EXP61103"
# EXP6 + 11 + 03 = EXP61103

print("Reverse engineering from your example:")
print(f"Expected pattern: {expected}")
print("This suggests:")
print("- Location digits: 11")
print("- Cross passage digits: 03")
print()

# Let me test different scenarios
test_cases = [
    {"cross": "B03", "location": "L11", "expected": "EXP61103"},
    {"cross": "3", "location": "11", "expected": "EXP61103"},
    {"cross": "03", "location": "11", "expected": "EXP61103"},
]

import re

for case in test_cases:
    cross = case["cross"]
    location = case["location"]
    expected = case["expected"]
    
    print(f"\nTesting: Cross='{cross}', Location='{location}'")
    
    # Method 1: Extract all digits, take last 2
    cross_digits = re.findall(r'\d', str(cross))
    location_digits = re.findall(r'\d', str(location))
    
    cross_2 = (cross_digits[-2:] if len(cross_digits) >= 2 else ['0'] * (2-len(cross_digits)) + cross_digits)
    location_2 = (location_digits[-2:] if len(location_digits) >= 2 else ['0'] * (2-len(location_digits)) + location_digits)
    
    pattern1 = f"EXP6{''.join(location_2)}{''.join(cross_2)}"
    print(f"  Method 1: {pattern1}")
    
    # Method 2: Extract numeric parts, take last 2 digits of each
    cross_nums = re.findall(r'\d+', str(cross))
    location_nums = re.findall(r'\d+', str(location))
    
    if cross_nums:
        cross_str = cross_nums[-1][-2:].zfill(2)
    else:
        cross_str = "00"
        
    if location_nums:
        location_str = location_nums[-1][-2:].zfill(2)
    else:
        location_str = "00"
    
    pattern2 = f"EXP6{location_str}{cross_str}"
    print(f"  Method 2: {pattern2}")
    
    print(f"  Expected: {expected}")
    print(f"  Match 1: {pattern1 == expected}")
    print(f"  Match 2: {pattern2 == expected}")
