<?php

require_once __DIR__ . '/vendor/autoload.php';

use Google\Client;
use Google\Service\Sheets;
use Blackmoon\SheetCrud\SheetCrud;

function getService(): Sheets
{
    $client = new Client();
    $client->setApplicationName('SheetCrud Complete Example');
    $client->setScopes([Sheets::SPREADSHEETS]);
    $client->setAuthConfig('path/to/your/credentials.json');
    return new Sheets($client);
}

try {
    $service = getService();
    $spreadsheetId = '107ngzC1BctThAtqIsZCk7k34KQTYUFxzI5Q3av3A7YY';
    $sheetCrud = new SheetCrud($service, $spreadsheetId);

    echo "1. Reading data:\n";
    $data = $sheetCrud->readFrom('Sheet1!A1:D5');
    print_r($data);

    echo "\n2. Appending data:\n";
    $newData = [['Appended', 'Row', date('Y-m-d H:i:s')]];
    $rowsAppended = $sheetCrud->appendTo('Sheet1!A:C', $newData);
    echo "Rows appended: $rowsAppended\n";

    echo "\n3. Inserting data:\n";
    $insertData = [['Inserted', 'Row', date('Y-m-d H:i:s')]];
    $rowsInserted = $sheetCrud->insertInto('Sheet1!A7:C7', $insertData);
    echo "Rows inserted: $rowsInserted\n";

    echo "\n4. Updating range:\n";
    $updateData = [['Updated', 'Cell']];
    $cellsUpdated = $sheetCrud->updateRange('Sheet1!A1:B1', $updateData);
    echo "Cells updated: $cellsUpdated\n";

    echo "\n5. Clearing range:\n";
    $cleared = $sheetCrud->clearRange('Sheet1!E1:F10');
    echo $cleared ? "Range cleared successfully\n" : "Failed to clear range\n";

    echo "\n6. Creating new sheet:\n";
    $newSheetName = 'NewSheet' . time();
    $newSheetId = $sheetCrud->createNewSheet($newSheetName);
    echo "New sheet created with ID: $newSheetId\n";

    echo "\n7. Renaming sheet:\n";
    $renamedSheetName = 'RenamedSheet' . time();
    $renamed = $sheetCrud->renameSheet($newSheetName, $renamedSheetName);
    echo $renamed ? "Sheet renamed successfully\n" : "Failed to rename sheet\n";

    echo "\n8. Deleting sheet:\n";
    $deleted = $sheetCrud->deleteSheet($renamedSheetName);
    echo $deleted ? "Sheet deleted successfully\n" : "Failed to delete sheet\n";

    echo "\n9. Duplicating sheet:\n";
    $duplicateSheetId = $sheetCrud->duplicateSheet('Sheet1', 'DuplicatedSheet');
    echo "Sheet duplicated with new ID: $duplicateSheetId\n";

    echo "\n10. Copying sheet to another spreadsheet:\n";
    $destinationSpreadsheetId = 'destination-spreadsheet-id';
    $copiedSheetId = $sheetCrud->copySheetToAnotherSpreadsheet('Sheet1', $destinationSpreadsheetId);
    echo "Sheet copied to another spreadsheet with ID: $copiedSheetId\n";

    echo "\n11. Getting columns:\n";
    $columns = $sheetCrud->getColumns('Sheet1!1:1');
    print_r($columns);

    echo "\n12. Inserting into columns:\n";
    $columnData = [['Column1', 'Column3'], ['Value1', 'Value3']];
    $rowsInserted = $sheetCrud->insertIntoColumns('Sheet1!A:C', $columnData);
    echo "Rows inserted: $rowsInserted\n";

    echo "\n13. Formatting range:\n";
    $format = ['backgroundColor' => ['red' => 0.8, 'green' => 0.8, 'blue' => 0.8], 'textFormat' => ['bold' => true]];
    $sheetCrud->formatRange('Sheet1!A1:B5', $format);
    echo "Range formatted\n";

    echo "\n14. Setting column width:\n";
    $widthSet = $sheetCrud->setColumnWidth('Sheet1', 0, 200); // Set column A width to 200 pixels
    echo $widthSet ? "Column width set successfully\n" : "Failed to set column width\n";

    echo "\n15. Setting row height:\n";
    $heightSet = $sheetCrud->setRowHeight('Sheet1', 0, 50); // Set row 1 height to 50 pixels
    echo $heightSet ? "Row height set successfully\n" : "Failed to set row height\n";

    echo "\n16. Protecting range:\n";
    $protectedRangeId = $sheetCrud->protectRange('Sheet1!A1:B5', ['user@example.com']);
    echo "Range protected with ID: $protectedRangeId\n";

    echo "\n17. Removing protection:\n";
    $protectionRemoved = $sheetCrud->removeProtection($protectedRangeId);
    echo $protectionRemoved ? "Protection removed successfully\n" : "Failed to remove protection\n";

    echo "\n18. Setting data validation:\n";
    $rule = ['condition' => ['type' => 'NUMBER_BETWEEN', 'values' => [['userEnteredValue' => '1'], ['userEnteredValue' => '10']]]];
    $validationSet = $sheetCrud->setDataValidation('Sheet1!C1:C10', $rule);
    echo $validationSet ? "Data validation set successfully\n" : "Failed to set data validation\n";

    echo "\n19. Adding named range:\n";
    $namedRangeId = $sheetCrud->addNamedRange('MyNamedRange', 'Sheet1!C1:D10');
    echo "Named range added with ID: $namedRangeId\n";

    echo "\n20. Deleting named range:\n";
    $namedRangeDeleted = $sheetCrud->deleteNamedRange($namedRangeId);
    echo $namedRangeDeleted ? "Named range deleted successfully\n" : "Failed to delete named range\n";

    echo "\n21. Adding chart:\n";
    $chartSpec = [
        'type' => 'BAR',
        'data' => ['sourceRange' => ['sources' => [['startRowIndex' => 0, 'startColumnIndex' => 0, 'endRowIndex' => 5, 'endColumnIndex' => 2]]]],
        'spec' => ['title' => 'My Chart']
    ];
    $chartId = $sheetCrud->addChart('Sheet1', $chartSpec);
    echo "Chart added with ID: $chartId\n";

    echo "\n22. Adding pivot table:\n";
    $pivotSpec = [
        'source' => ['sheetId' => 0, 'startRowIndex' => 0, 'startColumnIndex' => 0, 'endRowIndex' => 20, 'endColumnIndex' => 7],
        'rows' => [['sourceColumnOffset' => 0, 'showTotals' => true, 'sortOrder' => 'ASCENDING']],
        'columns' => [['sourceColumnOffset' => 1, 'showTotals' => true, 'sortOrder' => 'ASCENDING']],
        'values' => [['summarizeFunction' => 'SUM', 'sourceColumnOffset' => 2]]
    ];
    $pivotAdded = $sheetCrud->addPivotTable('Sheet1', 'PivotSheet', $pivotSpec);
    echo $pivotAdded ? "Pivot table added successfully\n" : "Failed to add pivot table\n";

    echo "\n23. Adding conditional format rule:\n";
    $rule = [
        'ranges' => [['sheetId' => 0, 'startRowIndex' => 0, 'endRowIndex' => 10, 'startColumnIndex' => 0, 'endColumnIndex' => 5]],
        'booleanRule' => [
            'condition' => ['type' => 'NUMBER_GREATER', 'values' => [['userEnteredValue' => '5']]],
            'format' => ['backgroundColor' => ['red' => 1, 'green' => 0, 'blue' => 0]]
        ]
    ];
    $ruleAdded = $sheetCrud->addConditionalFormatRule('Sheet1', $rule);
    echo $ruleAdded ? "Conditional format rule added successfully\n" : "Failed to add conditional format rule\n";

    echo "\n24. Sorting range:\n";
    $sortSpecs = [['dimensionIndex' => 0, 'sortOrder' => 'ASCENDING']];
    $sorted = $sheetCrud->sortRange('Sheet1!A1:C10', $sortSpecs);
    echo $sorted ? "Range sorted successfully\n" : "Failed to sort range\n";

    echo "\n25. Adding filter:\n";
    $filterAdded = $sheetCrud->addFilter('Sheet1!A1:C10');
    echo $filterAdded ? "Filter added successfully\n" : "Failed to add filter\n";

    echo "\n26. Merging cells:\n";
    $merged = $sheetCrud->mergeCells('Sheet1!D1:E2');
    echo $merged ? "Cells merged successfully\n" : "Failed to merge cells\n";

    echo "\n27. Unmerging cells:\n";
    $unmerged = $sheetCrud->unmergeCells('Sheet1!D1:E2');
    echo $unmerged ? "Cells unmerged successfully\n" : "Failed to unmerge cells\n";

    echo "\n28. Setting frozen rows:\n";
    $frozenRows = $sheetCrud->setFrozenRows('Sheet1', 1);
    echo $frozenRows ? "Frozen rows set successfully\n" : "Failed to set frozen rows\n";

    echo "\n29. Setting frozen columns:\n";
    $frozenColumns = $sheetCrud->setFrozenColumns('Sheet1', 1);
    echo $frozenColumns ? "Frozen columns set successfully\n" : "Failed to set frozen columns\n";

    echo "\n30. Batch update:\n";
    $batchOperations = [
        ['updateCells' => [
            'rows' => [['values' => [['userEnteredValue' => ['stringValue' => 'Batch Updated']]]]],
            'fields' => 'userEnteredValue',
            'range' => ['sheetId' => 0, 'startRowIndex' => 0, 'endRowIndex' => 1, 'startColumnIndex' => 0, 'endColumnIndex' => 1]
        ]]
    ];
    $sheetCrud->batchUpdate($batchOperations);
    echo "Batch update completed\n";

    echo "\nAll operations completed successfully!\n";

} catch (Exception $e) {
    echo "Error: " . $e->getMessage() . "\n";
    echo "In file: " . $e->getFile() . " on line " . $e->getLine() . "\n";
}