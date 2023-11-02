<?php

namespace Marcio1002\Work;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Excel
{
    public static function import(string $readFile, string $writeFile): void
    {
        try {
            static::validate($readFile);
            static::validate($writeFile);

            $data = static::readExcel($readFile);

            static::writeExcel($writeFile, $data);
        } catch (\Throwable $th) {
            echo $th->getMessage();
        }
    }

    private static function readExcel(string $file_excel): array
    {
        $excel_valid = [
            IOFactory::READER_ODS,
            IOFactory::READER_CSV,
            IOFactory::READER_XLS,
            IOFactory::READER_XLSX,
        ];

        $read_columns = ['J', 'K', 'M', 'B', 'A', 'C', 'N'];

        $spreadsheet = IOFactory::load(
            filename: $file_excel,
            flags: 0,
            readers: $excel_valid,
        );

        $worksheet = $spreadsheet->getSheet(0);

        $data = [];

        foreach ($worksheet->getRowIterator() as $row) {
            if ($row->getRowIndex() === 1) continue;

            foreach ($read_columns as $column) {
                if (!\key_exists($column, $data)) $data[$column] = [];

                $data[$column][] = $worksheet->getCell("$column{$row->getRowIndex()}")->getValue();
            }
        }

        return $data;
    }

    private static function writeExcel(string $file_excel, array $data): void
    {
        $write_columns = [
            'J' => 'A',
            'K' => 'B',
            'M' => 'C',
            'B' => 'D',
            'A' => 'E',
            'C' => 'F',
            'N' => 'H'
        ];

        $excel_valid = [
            IOFactory::READER_ODS,
            IOFactory::READER_CSV,
            IOFactory::READER_XLS,
            IOFactory::READER_XLSX,
        ];

        $spreadsheet = IOFactory::load(
            filename: $file_excel,
            flags: 0,
            readers: $excel_valid,
        );

        $worksheet = $spreadsheet->getSheet(1);

        foreach ($write_columns as $key_data => $column) {
            foreach ($data[$key_data] as $key_row => $row) {
                $key_row += 2;

                if (!\in_array($column, ['A', 'B', 'C', 'H'])) {
                    $worksheet->setCellValue("{$column}{$key_row}", $row);
                    continue;
                }

                // $date = join("-", array_reverse(explode("/", $data['J'][$key_row - 2])));

                $column === 'A'
                    ? $worksheet->setCellValue("{$column}{$key_row}", Date::PHPToExcel($row))
                    : $worksheet->setCellValue("{$column}{$key_row}", $row);

                static::formateRow($worksheet, $column, $key_row);
            }
        }

        static::formateColumnG($worksheet, $data);

        $rows = $worksheet->getHighestRow();
        $row_total = $rows +1;
        $worksheet->setCellValue("A$row_total", "Total");
        $worksheet->getStyle("A$row_total")->getFont()->setBold(true);
        $worksheet->setCellValue("H$row_total", "=SUM(H2:H$rows)");

        $writer = IOFactory::createWriter($spreadsheet, IOFactory::READER_XLSX);

        $exts = \join("|", $GLOBALS['extensions_valid']);
        $now = \date("d_m_Y__H-i-s");
        $file_excel_imported = \preg_replace(
            pattern: "/(?:[^\/])+(?:$exts)/",
            replacement: "{$now}.xlsx",
            subject: $file_excel
        );

        $writer->save($file_excel_imported);

        Std::write("Dados importado em \"$file_excel_imported\"");
        Std::write("Por favor verifique seu arquivo e check se está tudo OK");
    }

    private static function formateRow(Worksheet $worksheet, string $column, int $key_row): void
    {
        $style_col = $worksheet->getStyle("{$column}{$key_row}");
        $style_col->getFont()->setBold(false);
        $style_col->getFont()->setColor(new Color(Color::COLOR_BLACK));

        $column === 'A'
            ? $style_col->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY)
            : $style_col->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_TIME3);
    }

    private static function formateColumnG(Worksheet $worksheet, array $data): void
    {
        foreach ($data['J'] as $key_row => $_) {
            $key_row += 2;

            $worksheet->setCellValue("G{$key_row}", "");
            $styleG = $worksheet->getStyle("G{$key_row}");
            $styleG->getFont()->setBold(false);
            $styleG->getFont()->setColor(new Color(Color::COLOR_BLACK));
            $styleG->getFont()->setUnderline(Font::UNDERLINE_NONE);
        }
    }

    private static function validate(string $file): void
    {
        $fileinfo = new \SplFileInfo($file);

        if (!$fileinfo->isFile()) {
            Std::error("\"$file\" Não é um arquivo");
            exit;
        }

        if (!\in_array(($ext = $fileinfo->getExtension()), $GLOBALS['extensions_valid'])) {
            Std::error("\"{$ext}\" é um formato inválido. Precisa ser: " . join(", ", $GLOBALS['extensions_valid']));
            exit;
        }

        if (!$fileinfo->isReadable()) {
            Std::error('Esse arquivo não pode ser lido');
            exit;
        }
    }
}
