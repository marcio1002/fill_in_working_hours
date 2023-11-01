<?php

require_once __DIR__ . "/vendor/autoload.php";

use Marcio1002\Work\Excel;
use Marcio1002\Work\Std;


$GLOBALS['extensions_valid'] = ['csv', 'xlsx', 'xls', 'ods'];

Std::write('Digite o diretório completo do arquivo excel do Clockfy: ');
$path_excel = Std::read();

Std::write('Agora, digite o diretório completo que contém o modelo excel de "Relatório de Horas"');
$path_excel_hours = Std::read();

Excel::import($path_excel, $path_excel_hours);