<?php
use Google\Service\Sheets;
use Blackmoon\SheetCrud\SheetOperations;

require_once __DIR__ . '/./vendor/autoload.php';

function getService(): Sheets
{
    $client = new Google\Client();
    $client->setApplicationName('Google Sheets API CRUD');
    $client->setScopes(Google_Service_Sheets::SPREADSHEETS);
    $client->setAuthConfig('credentials.json');
    $client->setAccessType('offline');
    return new Google\Service\Sheets($client);
}

try {
    $service = getService();
    $spreadsheetId = '107ngzC1BctThAtqIsZCk7k34KQTYUFxzI5Q3av3A7YY';
    $sheetName = 'Sheet1';

    $sheet1 = new SheetOperations($service, $spreadsheetId);

    echo "Available methods:\n";
    print_r(get_class_methods($sheet1));

    // Read data
    echo "\nAttempting to read data...\n";
    $range = $sheetName . '!A1:D10';
    echo "Using range: $range\n";
    $result = $sheet1->readFrom($range);
    print_r($result);

    // Insert data
    $values = [
        ['New Column 1', 'New Column 2'],
        ['New Value 1', 'New Value 2']
    ];
    $insertRange = $sheetName . '!A:B';
    echo "\nInserting data to range: $insertRange\n";
    $insertedRows = $sheet1->insertInto($insertRange, $values);
    echo "$insertedRows rows inserted\n";

    // Get columns
    $columnRange = $sheetName . '!1:1';
    echo "\nGetting columns from range: $columnRange\n";
    $columns = $sheet1->getColumns($columnRange);
    echo "Columns:\n";
    print_r($columns);

    // Insert into specific columns
    $columnValues = [
        ['Column 1', 'Column 2'],
        ['Specific Value 1', 'Specific Value 2']
    ];
    $insertColumnRange = $sheetName . '!A:B';
    echo "\nInserting into specific columns, range: $insertColumnRange\n";
    $insertedRows = $sheet1->insertIntoColumns($insertColumnRange, $columnValues);
    echo "$insertedRows rows inserted into specific columns\n";

    // Format range
    $formatRange = $sheetName . '!A1:B2';
    echo "\nFormatting range: $formatRange\n";
    $format = [
        'backgroundColor' => ['red' => 0.8, 'green' => 0.8, 'blue' => 0.8],
        'textFormat' => ['bold' => true]
    ];
    $sheet1->formatRange($formatRange, $format);
    echo "Formatted range $formatRange\n";

    // Clear range
    $clearRange = $sheetName . '!D1:E5';
    echo "\nClearing range: $clearRange\n";
    $cleared = $sheet1->clearRange($clearRange);
    echo $cleared ? "Range $clearRange cleared\n" : "Failed to clear range $clearRange\n";

    // Update range
    $updateRange = $sheetName . '!A1:B2';
    echo "\nUpdating range: $updateRange\n";
    $updateValues = [
        ['Updated Value 1', 'Updated Value 2'],
        ['Updated Value 3', 'Updated Value 4']
    ];
    $updatedCells = $sheet1->updateRange($updateRange, $updateValues);
    echo "$updatedCells cells updated in range $updateRange\n";

     // Create a new sheet
    $newSheetName = 'NewSheet' . time(); // Appending timestamp to ensure unique name
    echo "\nCreating a new sheet named: $newSheetName\n";
    $newSheetId = $sheet1->createNewSheet($newSheetName);
    echo "New sheet created with ID: $newSheetId\n";

    // Rename the newly created sheet
    $renamedSheetName = 'RenamedSheet' . time();
    echo "\nRenaming sheet '$newSheetName' to '$renamedSheetName'\n";
    $renamed = $sheet1->renameSheet($newSheetName, $renamedSheetName);
    if ($renamed) {
        echo "Sheet successfully renamed to $renamedSheetName\n";
    } else {
        echo "Failed to rename sheet\n";
    }

} catch (\Exception $e) {
    echo "Error: " . $e->getMessage() . "\n";
    echo "In file: " . $e->getFile() . " on line " . $e->getLine() . "\n";
}