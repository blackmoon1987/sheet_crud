<?php

namespace Blackmoon\SheetCrud;


use Google\Service\Sheets;
use Google_Service_Sheets_ValueRange;
use Google_Service_Sheets_ClearValuesRequest;
use Google_Service_Sheets_BatchUpdateSpreadsheetRequest;
use Google_Service_Sheets_Request;
use Google_Service_Sheets_CellFormat;
use Google_Service_Sheets_CellData;
use Google_Service_Sheets_ExtendedValue;
use Exception;

class SheetOperations
{
    private Sheets $service;
    private string $spreadsheetId;

    public function __construct(Sheets $service, string $spreadsheetId)
    {
        $this->service = $service;
        $this->spreadsheetId = $spreadsheetId;
    }

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

    public function insertInto(string $range, array $values, string $insertAs = 'RAW'): int
    {
        return $this->appendTo($range, $values, $insertAs);
    }
    
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

    private function getLastRowIndex(string $range): int
    {
        try {
            $response = $this->service->spreadsheets_values->get($this->spreadsheetId, $range);
            return count($response->getValues() ?? []);
        } catch (Exception $e) {
            throw new Exception("Failed to get last row index: " . $e->getMessage());
        }
    }

    private function validateRange(string $range): void
    {
        // This regex allows for:
        // - Sheet names with spaces and special characters
        // - Open-ended column ranges like 'A:B'
        // - Row-only ranges like '1:1'
        // - Standard A1 notation ranges
        if (!preg_match('/^.+!([A-Z]+(\d+)?:)?[A-Z]+(\d+)?$|^.+!\d+:\d+$|^[A-Z]+\d+:[A-Z]+\d+$|^[A-Z]+:[A-Z]+$|^\d+:\d+$/', $range)) {
            throw new Exception("Invalid range format: $range");
        }
    }

    private function validateValues(array $values): void
    {
        if (empty($values) || !is_array($values[0])) {
            throw new Exception("Invalid values format. Expected 2D array.");
        }
    }

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
     * Renames an existing sheet.
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
     * Creates a new sheet in the spreadsheet.
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
}