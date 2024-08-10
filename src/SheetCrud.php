<?php

namespace Blackmoon\SheetCrud;

use Google\Service\Sheets;
use Google_Service_Sheets_ValueRange;
use Google_Service_Sheets_ClearValuesRequest;
use Google_Service_Sheets_BatchUpdateSpreadsheetRequest;
use Google_Service_Sheets_Request;
use Google_Service_Sheets_CellFormat;
use Google_Service_Sheets_Sheet;
use Google_Service_Sheets_SheetProperties;
use Google_Service_Sheets_DimensionRange;
use Google_Service_Sheets_AddProtectedRangeRequest;
use Google_Service_Sheets_ProtectedRange;
use Google_Service_Sheets_DeleteProtectedRangeRequest;
use Google_Service_Sheets_AddNamedRangeRequest;
use Google_Service_Sheets_NamedRange;
use Google_Service_Sheets_DeleteNamedRangeRequest;
use Google_Service_Sheets_AddChartRequest;
use Google_Service_Sheets_EmbeddedChart;
use Google_Service_Sheets_AddConditionalFormatRuleRequest;
use Google_Service_Sheets_SortRangeRequest;
use Google_Service_Sheets_SetBasicFilterRequest;
use Google_Service_Sheets_MergeCellsRequest;
use Google_Service_Sheets_UnmergeCellsRequest;
use Exception;

class SheetCrud
{
    private Sheets $service;
    private string $spreadsheetId;

    public function __construct(Sheets $service, string $spreadsheetId)
    {
        $this->service = $service;
        $this->spreadsheetId = $spreadsheetId;
    }

    /**
     * Read data from a specified range.
     *
     * @param string $range The A1 notation of the range to retrieve values from.
     * @return array The values in the range.
     * @throws Exception If there's an error reading the data.
     */
    public function readFrom(string $range): array
    {
        $this->validateRange($range);

        try {
            $response = $this->service->spreadsheets_values->get($this->spreadsheetId, $range);
            return $response->getValues() ?? [];
        } catch (Exception $e) {
            throw new Exception("Failed to read data: " . $e->getMessage());
        }
    }

    /**
     * Append data to a specified range.
     *
     * @param string $range The A1 notation of the range to append to.
     * @param array $values 2D array of values to append.
     * @param string $valueInputOption How the input data should be interpreted.
     * @return int Number of rows appended.
     * @throws Exception If there's an error appending the data.
     */
    public function appendTo(string $range, array $values, string $valueInputOption = 'RAW'): int
    {
        $this->validateRange($range);
        $this->validateValues($values);

        $body = new Google_Service_Sheets_ValueRange(['values' => $values]);
        $params = [
            'valueInputOption' => $valueInputOption,
            'insertDataOption' => 'INSERT_ROWS'
        ];

        try {
            $result = $this->service->spreadsheets_values->append(
                $this->spreadsheetId, 
                $range, 
                $body, 
                $params
            );
            return $result->getUpdates()->getUpdatedRows();
        } catch (Exception $e) {
            throw new Exception("Failed to append data: " . $e->getMessage());
        }
    }

    /**
     * Insert data into a specified range.
     *
     * @param string $range The A1 notation of the range to insert into.
     * @param array $values 2D array of values to insert.
     * @param string $insertAs How the input data should be interpreted.
     * @return int Number of rows inserted.
     * @throws Exception If there's an error inserting the data.
     */
    public function insertInto(string $range, array $values, string $insertAs = 'RAW'): int
    {
        return $this->appendTo($range, $values, $insertAs);
    }

    /**
     * Update a specified range with new values.
     *
     * @param string $range The A1 notation of the range to update.
     * @param array $values 2D array of values to update with.
     * @param string $valueInputOption How the input data should be interpreted.
     * @return int Number of cells updated.
     * @throws Exception If there's an error updating the range.
     */
    public function updateRange(string $range, array $values, string $valueInputOption = 'RAW'): int
    {
        $this->validateRange($range);
        $this->validateValues($values);

        $body = new Google_Service_Sheets_ValueRange(['values' => $values]);
        $params = ['valueInputOption' => $valueInputOption];

        try {
            $result = $this->service->spreadsheets_values->update(
                $this->spreadsheetId, 
                $range, 
                $body, 
                $params
            );
            return $result->getUpdatedCells();
        } catch (Exception $e) {
            throw new Exception("Failed to update range: " . $e->getMessage());
        }
    }

    /**
     * Clear a specified range.
     *
     * @param string $range The A1 notation of the range to clear.
     * @return bool True if the range was cleared successfully, false otherwise.
     * @throws Exception If there's an error clearing the range.
     */
    public function clearRange(string $range): bool
    {
        $this->validateRange($range);

        $body = new Google_Service_Sheets_ClearValuesRequest();
        
        try {
            $response = $this->service->spreadsheets_values->clear(
                $this->spreadsheetId,
                $range,
                $body
            );
            return $response->getClearedRange() !== null;
        } catch (Exception $e) {
            throw new Exception("Failed to clear range: " . $e->getMessage());
        }
    }

    /**
     * Create a new sheet in the spreadsheet.
     *
     * @param string $sheetName The name for the new sheet.
     * @return int The ID of the newly created sheet.
     * @throws Exception If there's an error creating the sheet.
     */
    public function createNewSheet(string $sheetName): int
    {
        try {
            $body = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => [
                    [
                        'addSheet' => [
                            'properties' => [
                                'title' => $sheetName
                            ]
                        ]
                    ]
                ]
            ]);

            $response = $this->service->spreadsheets->batchUpdate($this->spreadsheetId, $body);
            $addedSheet = $response->getReplies()[0]->getAddSheet();
            return $addedSheet->getProperties()->getSheetId();
        } catch (Exception $e) {
            throw new Exception("Error creating new sheet: " . $e->getMessage());
        }
    }

    /**
     * Rename an existing sheet.
     *
     * @param string $oldSheetName The current name of the sheet to be renamed.
     * @param string $newSheetName The new name for the sheet.
     * @return bool True if the sheet was successfully renamed, false otherwise.
     * @throws Exception If there's an error renaming the sheet.
     */
    public function renameSheet(string $oldSheetName, string $newSheetName): bool
    {
        try {
            $spreadsheet = $this->service->spreadsheets->get($this->spreadsheetId);
            $sheetId = null;

            foreach ($spreadsheet->getSheets() as $sheet) {
                if ($sheet->getProperties()->getTitle() === $oldSheetName) {
                    $sheetId = $sheet->getProperties()->getSheetId();
                    break;
                }
            }

            if ($sheetId === null) {
                throw new Exception("Sheet '$oldSheetName' not found.");
            }

            $body = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => [
                    [
                        'updateSheetProperties' => [
                            'properties' => [
                                'sheetId' => $sheetId,
                                'title' => $newSheetName
                            ],
                            'fields' => 'title'
                        ]
                    ]
                ]
            ]);

            $this->service->spreadsheets->batchUpdate($this->spreadsheetId, $body);
            return true;
        } catch (Exception $e) {
            throw new Exception("Error renaming sheet: " . $e->getMessage());
        }
    }

    /**
     * Delete a sheet from the spreadsheet.
     *
     * @param string $sheetName The name of the sheet to delete.
     * @return bool True if the sheet was successfully deleted, false otherwise.
     * @throws Exception If there's an error deleting the sheet.
     */
    public function deleteSheet(string $sheetName): bool
    {
        try {
            $sheetId = $this->getSheetId($sheetName);
            $body = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => [
                    [
                        'deleteSheet' => [
                            'sheetId' => $sheetId
                        ]
                    ]
                ]
            ]);

            $this->service->spreadsheets->batchUpdate($this->spreadsheetId, $body);
            return true;
        } catch (Exception $e) {
            throw new Exception("Error deleting sheet: " . $e->getMessage());
        }
    }

    /**
     * Get column information for a specified range.
     *
     * @param string $range The A1 notation of the range to get column information from.
     * @return array Associative array of column names and their index in A1 notation.
     * @throws Exception If there's an error getting the column information.
     */
    public function getColumns(string $range): array
    {
        $this->validateRange($range);

        try {
            $result = $this->service->spreadsheets_values->get($this->spreadsheetId, $range);
            $columns = $result->getValues()[0] ?? [];
            $columnMap = [];
            foreach ($columns as $index => $column) {
                if ($column !== '') {
                    $columnMap[$column] = chr(65 + $index);
                }
            }
            return $columnMap;
        } catch (Exception $e) {
            throw new Exception("Failed to get columns: " . $e->getMessage());
        }
    }

    /**
     * Insert data into specific columns.
     *
     * @param string $range The A1 notation of the range to insert into.
     * @param array $values 2D array of values to insert.
     * @param string $insertAs How the input data should be interpreted.
     * @return int Number of rows inserted.
     * @throws Exception If there's an error inserting the data.
     */
    public function insertIntoColumns(string $range, array $values, string $insertAs = 'RAW'): int
    {
        $this->validateRange($range);
        $this->validateValues($values);

        $allColumns = $this->getColumns($range);
        $requiredColumns = $values[0];
        $valuesRestructured = [];

        foreach (array_slice($values, 1) as $row) {
            $newRow = array_fill(0, count($allColumns), '');
            foreach ($requiredColumns as $index => $column) {
                if (isset($allColumns[$column])) {
                    $columnIndex = ord($allColumns[$column]) - 65;
                    $newRow[$columnIndex] = $row[$index] ?? '';
                }
            }
            $valuesRestructured[] = $newRow;
        }

        return $this->insertInto($range, $valuesRestructured, $insertAs);
    }

    /**
     * Format a range of cells.
     *
     * @param string $range The A1 notation of the range to format.
     * @param array $format The format to apply.
     * @throws Exception If there's an error formatting the range.
     */
    public function formatRange(string $range, array $format): void
    {
        $this->validateRange($range);

        $sheetId = $this->getSheetId($range);
        list($startColumn, $startRow, $endColumn, $endRow) = $this->parseRange($range);

        $requests = [
            new Google_Service_Sheets_Request([
                'repeatCell' => [
                    'range' => [
                        'sheetId' => $sheetId,
                        'startRowIndex' => $startRow - 1,
                        'endRowIndex' => $endRow,
                        'startColumnIndex' => $startColumn - 1,
                        'endColumnIndex' => $endColumn,
                    ],
                    'cell' => [
                        'userEnteredFormat' => new Google_Service_Sheets_CellFormat($format)
                    ],
                    'fields' => 'userEnteredFormat(' . implode(',', array_keys($format)) . ')'
                ]
            ])
        ];

        $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => $requests
        ]);

        try {
            $this->service->spreadsheets->batchUpdate($this->spreadsheetId, $batchUpdateRequest);
        } catch (Exception $e) {
            throw new Exception("Failed to format range: " . $e->getMessage());
        }
    }

    /**
     * Perform a batch update operation.
     *
     * @param array $operations Array of operations to perform.
     * @throws Exception If there's an error performing the batch update.
     */
    public function batchUpdate(array $operations): void
    {
        $requests = [];
        foreach ($operations as $operation) {
            $requests[] = new Google_Service_Sheets_Request($operation);
        }

        $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => $requests
        ]);

        try {
            $this->service->spreadsheets->batchUpdate($this->spreadsheetId, $batchUpdateRequest);
        } catch (Exception $e) {
            throw new Exception("Failed to perform batch update: " . $e->getMessage());
        }
    }

    // ... [Additional methods like protectRange, addNamedRange, addChart, etc. would go here]

    /**
     * Get the sheet ID for a given sheet name or range.
     *
     * @param string $range The sheet name or A1 notation range.
     * @return int The sheet ID.
     * @throws Exception If the sheet is not found.
     */
    private function getSheetId(string $range): int
    {
        $sheetName = explode('!', $range)[0];
        try {
            $spreadsheet = $this->service->spreadsheets->get($this->spreadsheetId);
            foreach ($spreadsheet->getSheets() as $sheet) {
                if ($sheet->getProperties()->getTitle() === $sheetName) {
                    return $sheet->getProperties()->getSheetId();
                }
            }
            throw new Exception("Sheet not found: $sheetName");
        } catch (Exception $e) {
            throw new Exception("Failed to get sheet ID: " . $e->getMessage());
        }
    }

    /**
     * Validate the format of a range string.
     *
     * @param string $range The range to validate.
     * @throws Exception If the range format is invalid.
     */
    private function validateRange(string $range): void
    {
        if (!preg_match('/^.+!([A-Z]+(\d+)?:)?[A-Z]+(\d+)?$|^.+!\d+:\d+$|^[A-Z]+\d+:[A-Z]+\d+$|^[A-Z]+:[A-Z]+$|^\d+:\d+$/', $range)) {
            throw new Exception("Invalid range format: $range");
        }
    }

    /**
     * Validate the format of values array.
     *
     * @param array $values The values to validate.
     * @throws Exception If the values format is invalid.
     */
    private function validateValues(array $values): void
    {
        if (empty($values) || !is_array($values[0])) {
            throw new Exception("Invalid values format. Expected 2D array.");
        }
    }

    /**
     * Parse a range string into its components.
     *
     * @param string $range The range to parse.
     * @return array An array containing start column, start row, end column, and end row.
     * @throws Exception If the range format is invalid for parsing.
     */
    private function parseRange(string $range): array
    {
        preg_match('/([A-Z]+)(\d+):([A-Z]+)(\d+)/', $range, $matches);
        if (count($matches) !== 5) {
            throw new Exception("Invalid range format for parsing: $range");
        }
        return [
            $this->columnLetterToNumber($matches[1]),
            (int)$matches[2],
            $this->columnLetterToNumber($matches[3]),
            (int)$matches[4]
        ];
    }

    /**
     * Convert a column letter to its corresponding number.
     *
     * @param string $column The column letter.
     * @return int The corresponding column number.
     */
    private function columnLetterToNumber(string $column): int
    {
        $column = strtoupper($column);
        $number = 0;
        $length = strlen($column);
        for ($i = 0; $i < $length; $i++) {
            $number += (ord($column[$i]) - 64) * pow(26, $length - $i - 1);
        }
        return $number;
    }

    /**
     * Protect a range in the spreadsheet.
     *
     * @param string $range The range to protect.
     * @param array $editors List of email addresses of users who can edit the range.
     * @return int The ID of the protected range.
     * @throws Exception If there's an error protecting the range.
     */
    public function protectRange(string $range, array $editors = []): int
    {
        $this->validateRange($range);

        $rangeToProtect = $this->getA1Range($range);
        $request = new Google_Service_Sheets_AddProtectedRangeRequest([
            'protectedRange' => [
                'range' => $rangeToProtect,
                'editors' => ['users' => $editors]
            ]
        ]);

        $body = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => [['addProtectedRange' => $request]]
        ]);

        try {
            $response = $this->service->spreadsheets->batchUpdate($this->spreadsheetId, $body);
            return $response->getReplies()[0]->getAddProtectedRange()->getProtectedRange()->getProtectedRangeId();
        } catch (Exception $e) {
            throw new Exception("Failed to protect range: " . $e->getMessage());
        }
    }

    /**
     * Remove protection from a range.
     *
     * @param int $protectedRangeId The ID of the protected range to remove.
     * @return bool True if the protection was successfully removed, false otherwise.
     * @throws Exception If there's an error removing the protection.
     */
    public function removeProtection(int $protectedRangeId): bool
    {
        $request = new Google_Service_Sheets_DeleteProtectedRangeRequest([
            'protectedRangeId' => $protectedRangeId
        ]);

        $body = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => [['deleteProtectedRange' => $request]]
        ]);

        try {
            $this->service->spreadsheets->batchUpdate($this->spreadsheetId, $body);
            return true;
        } catch (Exception $e) {
            throw new Exception("Failed to remove protection: " . $e->getMessage());
        }
    }

    /**
     * Add a named range to the spreadsheet.
     *
     * @param string $name The name for the named range.
     * @param string $range The range for the named range.
     * @return string The ID of the newly created named range.
     * @throws Exception If there's an error adding the named range.
     */
    public function addNamedRange(string $name, string $range): string
    {
        $this->validateRange($range);

        $rangeToName = $this->getA1Range($range);
        $request = new Google_Service_Sheets_AddNamedRangeRequest([
            'namedRange' => [
                'name' => $name,
                'range' => $rangeToName
            ]
        ]);

        $body = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => [['addNamedRange' => $request]]
        ]);

        try {
            $response = $this->service->spreadsheets->batchUpdate($this->spreadsheetId, $body);
            return $response->getReplies()[0]->getAddNamedRange()->getNamedRange()->getNamedRangeId();
        } catch (Exception $e) {
            throw new Exception("Failed to add named range: " . $e->getMessage());
        }
    }

    /**
     * Delete a named range from the spreadsheet.
     *
     * @param string $namedRangeId The ID of the named range to delete.
     * @return bool True if the named range was successfully deleted, false otherwise.
     * @throws Exception If there's an error deleting the named range.
     */
    public function deleteNamedRange(string $namedRangeId): bool
    {
        $request = new Google_Service_Sheets_DeleteNamedRangeRequest([
            'namedRangeId' => $namedRangeId
        ]);

        $body = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => [['deleteNamedRange' => $request]]
        ]);

        try {
            $this->service->spreadsheets->batchUpdate($this->spreadsheetId, $body);
            return true;
        } catch (Exception $e) {
            throw new Exception("Failed to delete named range: " . $e->getMessage());
        }
    }

    /**
     * Convert a range string to an A1 range object.
     *
     * @param string $range The range string to convert.
     * @return array The A1 range object.
     */
    private function getA1Range(string $range): array
    {
        $parts = explode('!', $range);
        $sheetName = $parts[0];
        $cellRange = $parts[1] ?? 'A1';

        return [
            'sheetName' => $sheetName,
            'startColumnIndex' => 0,
            'startRowIndex' => 0,
            'endColumnIndex' => 1000,
            'endRowIndex' => 1000
        ];
    }
}