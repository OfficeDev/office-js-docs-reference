#!/bin/bash
# Test script to verify "Used By" feature implementation

echo "=== Testing Used By Feature ==="
echo ""

# Test 1: Check that ContextInformation has Used By section
echo "Test 1: Office.ContextInformation should show 'Used by' section..."
if grep -q "#### Used by" yaml/office/office/office.contextinformation.yml; then
    echo "✓ PASS: Used By section found"
    grep -A 3 "#### Used by" yaml/office/office/office.contextinformation.yml | head -4
else
    echo "✗ FAIL: Used By section not found"
fi
echo ""

# Test 2: Check that EventType has cross-package references
echo "Test 2: Office.EventType should show cross-package references..."
if grep -q "**Office**" yaml/office/office/office.eventtype.yml && grep -q "**Outlook**" yaml/office/office/office.eventtype.yml; then
    echo "✓ PASS: Cross-package references found"
else
    echo "✗ FAIL: Cross-package references not found"
fi
echo ""

# Test 3: Check that multiple references are listed
echo "Test 3: Excel.AggregationFunction should have multiple references..."
ref_count=$(grep -A 20 "#### Used by" yaml/excel/excel/excel.aggregationfunction.yml | grep -c "xref uid")
if [ "$ref_count" -gt 1 ]; then
    echo "✓ PASS: Found $ref_count references"
else
    echo "✗ FAIL: Expected multiple references, found $ref_count"
fi
echo ""

# Test 4: Check that Used By appears before Examples
echo "Test 4: Used By should appear before Examples..."
if grep -B 5 "#### Examples" yaml/outlook/outlook/office.diagnostics.yml | grep -q "#### Used by"; then
    echo "✓ PASS: Used By appears before Examples"
else
    echo "✗ FAIL: Used By does not appear before Examples"
fi
echo ""

# Test 5: Count total files with Used By sections
echo "Test 5: Counting files with Used By sections..."
total_count=$(find yaml -name "*.yml" -exec grep -l "#### Used by" {} \; | wc -l)
echo "✓ Found $total_count files with 'Used By' sections"
echo ""

echo "=== All Tests Complete ==="
