<?php

require_once __DIR__ . "/vendor/autoload.php";

use Marcio1002\Work\Excel;
use Marcio1002\Work\Std;

$GLOBALS['extensions_valid'] = ['csv', 'xlsx', 'xls', 'ods'];

Std::write('Digite o diretório completo do arquivo excel do Clockfy: ');
// $path_excel_import = Std::read();
$path_excel_import = "/home/marcio/Documentos/Clockify_Relatório_De_Tempo_Detalhado_01_03_2024-31_03_2024.xlsx";

Std::write('Agora, digite o diretório completo que contém o modelo excel de "Relatório de Horas"');
// $path_excel_export = Std::read();
$path_excel_export = "/home/marcio/Documentos/modelo-gebit.xlsx";

Excel::import(
    read_file: $path_excel_import,
    write_file: $path_excel_export
);