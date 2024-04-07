<?php

namespace Marcio1002\Work;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Shared\Date;

class Excel
{
    /**
     * 
     * 
     * @var array<string, string> $write_columns 
     */
    private array $write_columns  = [
        'J' => 'A',
        'K' => 'B',
        'M' => 'C',
        'B' => 'D',
        'A' => 'E',
        'C' => 'F',
        'N' => 'H'
    ];

    /** 
     * Arquivos válidos para importação e exportação
     * 
     * @var array<string> $excel_valid
     */
    private array $excel_valid = [
        IOFactory::READER_ODS,
        IOFactory::READER_CSV,
        IOFactory::READER_XLS,
        IOFactory::READER_XLSX,
    ];

    /**
     * Colunas que serão lidas no arquivo de importação
     *
     * @var array<string> $read_columns 
     */
    private array $read_columns = ['J', 'K', 'M', 'B', 'A', 'C', 'N'];


    private function __construct()
    {
    }

    public static function import(string $read_file, string $write_file): void
    {
        try {
            $self_instance  = new static();

            $self_instance->validate($read_file);
            $self_instance->validate($write_file);

            $data = $self_instance->readExcel($read_file);

            $self_instance->writeExcel($write_file, $data);
        } catch (\Throwable $th) {
            echo $th->getMessage();
        }
    }

    private function readExcel(string $file_excel): array
    {
        $spreadsheet = IOFactory::load(
            filename: $file_excel,
            flags: 0,
            readers: $this->excel_valid,
        );

        $worksheet = $spreadsheet->getSheet(0);

        $data = [];

        foreach ($worksheet->getRowIterator() as $row) {
            if ($row->getRowIndex() === 1) continue;

            foreach ($this->read_columns as $column) {
                if (!\key_exists($column, $data)) $data[$column] = [];

                $data[$column][] = $worksheet->getCell("$column{$row->getRowIndex()}")->getValue();
            }
        }

        return $data;
    }

    private function writeExcel(string $file_excel, array $data): void
    {
        $spreadsheet = IOFactory::load(
            filename: $file_excel,
            flags: 0,
            readers: $this->excel_valid,
        );

        $worksheet = $spreadsheet->getSheet(1);

        foreach ($this->write_columns as $key_data => $column) {
            foreach ($data[$key_data] as $key_row => $row) {
                if (empty($row?->getPlainText())) continue;

                $key_row += 2;

                $style_col = $worksheet->getStyle("{$column}{$key_row}");
                $style_col->getFont()->setBold(false);
                $style_col->getFont()->setColor(new Color(Color::COLOR_BLACK));

                if (!\in_array($column, ['A', 'B', 'C', 'H'])) {
                    $worksheet->setCellValue("{$column}{$key_row}", $row);
                    continue;
                }

                if ($column === 'A') {
                    $row = \preg_replace("/^(\d{4})-(\d{2})-(\d{2})$/", "$3\/$2\/$1", $row->getPlainText());
                    $row = \join(",", \explode('/', $row));
                    $row = "=DATE($row)";
                } else {
                    $row  = \preg_replace("/^(\d{2}):(\d{2})$/", "$1:$2:00", $row->getPlainText());
                    $row = \join(",", \explode(':', $row));
                    $row = "=TIME($row)";
                }

                $worksheet->setCellValueExplicit("{$column}{$key_row}", $row, DataType::TYPE_FORMULA);

                $this->formateRow($worksheet, $column, $key_row);
            }
        }

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

    private function formateRow(Worksheet $worksheet, string $column, int $key_row): void
    {
        $style_col = $worksheet->getStyle("{$column}{$key_row}");

        $column === 'A'
            ? $style_col->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY)
            : $style_col->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_TIME4);
    }

    private function validate(string $file): void
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
