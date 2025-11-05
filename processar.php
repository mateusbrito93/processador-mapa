<?php
// 1. Incluir o autoloader do Composer
require 'vendor/autoload.php';

// Aumenta o limite de memória por segurança
// (Se 512MB for o seu limite, vamos tentar 1GB)
ini_set('memory_limit', '1G'); 

// Define o fuso horário para garantir a data correta
date_default_timezone_set('America/Fortaleza');

use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

// Definições de Colunas
const COLUNA_FINAL_INDEX = 12; // L é a 12ª
const COLUNA_DOADOR_INDEX = 3;  // Doador é a 4ª (índice 3)
const COLUNA_DTNASC_INDEX = 4; // DT. NASC. é a 5ª (índice 4)
const COLUNA_TUBO_NOME = 'DOAÇAO TUBO NAT';

// Mapa de Cidades pelo código
const MAPA_CIDADES = [
    '04' => 'cajazeiras',
    '07' => 'catole',
    '05' => 'guarabira',
    '12' => 'itaporanga',
    '03' => 'patos',
    '09' => 'pianco',
    '06' => 'sousa'
];

// Lista de termos para filtrar
$termos_invalidos = [
    'NULO',
    'NULA',
    'CANCELADA',
    'CANCELADO',
    'INACESSIBILIDADE DE VEIA'
];

// Arrays de dados
$dados_validos = [];
$cabecalhos_array = [];
$nome_cidade = null; 

// Contadores para os avisos
$linhas_removidas_count = 0;
$termos_encontrados = [];

// Arrays para TODOS os erros de validação
$date_validation_errors = [];
$doacao_validation_errors = [];
$codigos_vistos = [];
$tubo_validation_errors = [];
$tubos_vistos = [];


// 2. Verificar se um arquivo foi enviado
if (isset($_FILES['arquivo_excel']) && $_FILES['arquivo_excel']['error'] == UPLOAD_ERR_OK) {

    $inputFileName = $_FILES['arquivo_excel']['tmp_name'];

    try {
        // 3. Carregar o arquivo Excel
        // (Otimização de Memória)
        
        // Identifica o tipo de arquivo (Xlsx, Xls, etc.)
        $fileType = IOFactory::identify($inputFileName);
        // Cria o leitor apropriado
        $reader = IOFactory::createReader($fileType);
        
        // Diz ao leitor para focar APENAS nos dados e ignorar estilos
        $reader->setReadDataOnly(true); 
        
        // Carrega o arquivo usando o leitor otimizado
        $spreadsheet = $reader->load($inputFileName);

        $sheet = $spreadsheet->getSheetByName('MAP MACR');

        if (!$sheet) {
            // Descarrega a planilha da memória antes de sair
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
            throw new Exception("Erro: A planilha 'MAP MACR' não foi encontrada.");
        }

        // 4. Encontrar a linha do cabeçalho
        $headerRow = null; 
        $startRow = null;  
        $coluna_tubo_index = -1;
        $highestRow = $sheet->getHighestDataRow();

        for ($row = 1; $row <= 20 && $row <= $highestRow; $row++) {
            $cellValue = (string) $sheet->getCell('A' . $row)->getValue();
            if (strtoupper(trim($cellValue)) === 'DOACAO HEMOVIDA') {
                $headerRow = $row;
                $startRow = $headerRow + 1; 
                break;
            }
        }

        if ($startRow === null) {
            $spreadsheet->disconnectWorksheets(); unset($spreadsheet);
            throw new Exception("Não foi possível encontrar a linha de cabeçalho 'DOACAO HEMOVIDA'.");
        }

        // 5. Pegar os cabeçalhos
        for ($col = 1; $col <= COLUNA_FINAL_INDEX; $col++) {
            $valor_cabecalho = trim((string) $sheet->getCell([$col, $headerRow])->getValue());
            $cabecalhos_array[] = $valor_cabecalho;
            
            if (strtoupper($valor_cabecalho) === COLUNA_TUBO_NOME) {
                $coluna_tubo_index = $col - 1; // -1 porque array é 0-indexed
            }
        }
        
        if ($coluna_tubo_index === -1) {
            $spreadsheet->disconnectWorksheets(); unset($spreadsheet);
            throw new Exception("Erro: A coluna '" . COLUNA_TUBO_NOME . "' não foi encontrada no cabeçalho.");
        }


        // 6. Iterar pelas linhas de dados
        for ($row = $startRow; $row <= $highestRow; $row++) {
            
            $linha_temporaria = [];
            $linha_contem_invalido = false;
            $termo_achado = '';
            
            // 7. VALIDAÇÃO 1: Campos Inválidos ("NULO", etc.)
            for ($col = 1; $col <= COLUNA_FINAL_INDEX; $col++) {
                $cellValue = $sheet->getCell([$col, $row])->getValue();
                $valor_checagem = strtoupper(trim((string)$cellValue));

                if (in_array($valor_checagem, $termos_invalidos)) {
                    $linha_contem_invalido = true;
                    $termo_achado = $valor_checagem; 
                    break; 
                }
                
                $linha_temporaria[] = $cellValue;
            }
            
            if ($linha_contem_invalido) {
                $linhas_removidas_count++;
                if (!in_array($termo_achado, $termos_encontrados)) {
                    $termos_encontrados[] = $termo_achado;
                }
                continue;
            }
            
            // 8. Se a linha for VÁLIDA, processa e transforma

            // (Req 1 e 2): Validação do CÓDIGO DE DOAÇÃO
            $doacao_code_raw = trim((string) $linha_temporaria[0]);
            $doacao_code = $doacao_code_raw;

            if (empty($doacao_code)) {
                continue; 
            }

            if (str_ends_with($doacao_code, ".0")) {
                $doacao_code = substr($doacao_code, 0, -2);
            }

            $erro_formato = false;
            if (!ctype_digit($doacao_code)) {
                $erro_formato = "não é composto apenas por números";
            } else if (strlen($doacao_code) != 11) {
                $erro_formato = "não possui 11 dígitos (encontrado: " . strlen($doacao_code) . ")";
            } else if (substr($doacao_code, 0, 1) != '6') {
                $erro_formato = "não começa com o dígito '6'";
            }
            
            if ($erro_formato) {
                $doacao_validation_errors[] = "Linha $row (Excel): Código '{$doacao_code_raw}' é inválido ($erro_formato).";
                continue; 
            }
            
            if (isset($codigos_vistos[$doacao_code])) {
                $linha_anterior = $codigos_vistos[$doacao_code];
                $doacao_validation_errors[] = "Linha $row (Excel): Código '$doacao_code' está duplicado (visto na linha $linha_anterior).";
                continue; 
            }
            $codigos_vistos[$doacao_code] = $row;


            // Define o nome da cidade na PRIMEIRA linha válida
            if ($nome_cidade === null) {
                $codigo_cidade = substr($doacao_code, 1, 2); // Pega 2º e 3º dígito
                
                if (isset(MAPA_CIDADES[$codigo_cidade])) {
                    $nome_cidade = MAPA_CIDADES[$codigo_cidade];
                } else {
                    $spreadsheet->disconnectWorksheets(); unset($spreadsheet);
                    throw new Exception("Linha $row (Excel): Código de cidade '{$codigo_cidade}' (extraído da doação '{$doacao_code}') não é reconhecido.");
                }
            }


            // (Req 3): Validação do CÓDIGO TUBO NAT
            $tubo_code_raw = trim((string) $linha_temporaria[$coluna_tubo_index]);
            $tubo_code = $tubo_code_raw;
            $erro_tubo = false;

            if (strlen($tubo_code) != 15) {
                $erro_tubo = "não possui 15 caracteres (encontrado: " . strlen($tubo_code) . ")";
            } else if (strtoupper(substr($tubo_code, 0, 1)) != 'B') {
                $erro_tubo = "não começa com a letra 'B'";
            }

            if ($erro_tubo) {
                $tubo_validation_errors[] = "Linha $row (Excel): Código Tubo '{$tubo_code_raw}' é inválido ($erro_tubo).";
                continue;
            }

            if (isset($tubos_vistos[$tubo_code])) {
                $linha_anterior = $tubos_vistos[$tubo_code];
                $tubo_validation_errors[] = "Linha $row (Excel): Código Tubo '$tubo_code' está duplicado (visto na linha $linha_anterior).";
                continue;
            }
            $tubos_vistos[$tubo_code] = $row;
            

            // REGRA DO HOSPITAL: Sobrescreve a coluna I (índice 8)
            $codigo_hospital_novo = substr($doacao_code, 1, 2); 
            $linha_temporaria[8] = $codigo_hospital_novo;
            
            // VALIDAÇÃO 3: Formato da Data de Nascimento
            $excel_date_serial = $linha_temporaria[COLUNA_DTNASC_INDEX];
            $linha_com_erro_de_data = false;

            if (is_numeric($excel_date_serial) && $excel_date_serial > 1) {
                try {
                    $datetime_obj = Date::excelToDateTimeObject($excel_date_serial);
                    $linha_temporaria[COLUNA_DTNASC_INDEX] = $datetime_obj->format('d/m/Y');
                } catch (Exception $dateEx) {
                    $linha_com_erro_de_data = true; 
                }
            } else if (!empty(trim((string)$excel_date_serial))) {
                $linha_com_erro_de_data = true;
            }
            
            if ($linha_com_erro_de_data) {
                $date_validation_errors[] = "Linha $row (Doação: $doacao_code) tem data inválida: '$excel_date_serial'";
            } else {
                $dados_validos[] = $linha_temporaria;
            }
        } // Fim do loop das linhas
        
        // Libera a memória do Excel o mais rápido possível
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
        unset($reader);


        // 7. VERIFICAR ERROS CRÍTICOS
        $all_critical_errors = array_merge($doacao_validation_errors, $tubo_validation_errors, $date_validation_errors);

        if (count($all_critical_errors) > 0) {
            $erros_para_mostrar = array_slice($all_critical_errors, 0, 10);
            $mensagem_erro = "Processamento interrompido. " . count($all_critical_errors) . " erro(s) crítico(s) encontrado(s):";
            $mensagem_erro .= " " . implode('; ', $erros_para_mostrar);
            
            throw new Exception($mensagem_erro);
        }

        // 8. Se não tiver dados válidos
        if (count($dados_validos) < 1) {
            throw new Exception("Nenhum dado válido encontrado após a filtragem.");
        }
        
        // 9. Gerar o arquivo CSV/TXT na memória
        $handle = fopen('php://memory', 'w');
        $delimiter = ';';
        $enclosure = '"';

        // Escrever cabeçalho
        fwrite($handle, implode($delimiter, $cabecalhos_array) . "\r\n"); 

        // Escrever dados
        foreach ($dados_validos as $linha) {
            $linha_para_escrever = [];
            foreach ($linha as $col_index => $valor) {
                $valor_str = (string) $valor;
                if ($col_index == COLUNA_DOADOR_INDEX) {
                    $linha_para_escrever[] = $valor_str;
                } else {
                    $valor_escapado = str_replace($enclosure, $enclosure . $enclosure, $valor_str);
                    if (preg_match("/[{$delimiter}{$enclosure} \t\n\r]/", $valor_escapado)) {
                        $linha_para_escrever[] = $enclosure . $valor_escapado . $enclosure;
                    } else {
                        $linha_para_escrever[] = $valor_escapado;
                    }
                }
            }
            fwrite($handle, implode($delimiter, $linha_para_escrever) . "\r\n");
        }

        // 10. Enviar o arquivo final para o usuário
        rewind($handle);
        $csv_content = stream_get_contents($handle);
        fclose($handle);

        $exposed_headers = ['Content-Disposition']; 

        if ($linhas_removidas_count > 0) {
            $mensagem_aviso = "Processamento concluído. ";
            $mensagem_aviso .= "$linhas_removidas_count linha(s) foram removidas por conterem os termos: " . implode(', ', $termos_encontrados) . ".";
            
            header("X-Warning-Message: " . rawurlencode($mensagem_aviso));
            $exposed_headers[] = 'X-Warning-Message'; 
        }
        
        header("Access-Control-Expose-Headers: " . implode(', ', $exposed_headers));

        $data_atual = date('d.m.Y');
        $nome_arquivo = "{$nome_cidade} {$data_atual}.txt";

        header('Content-Type: text/plain'); 
        header('Content-Disposition: attachment; filename="' . $nome_arquivo . '"');
        header('Content-Length: ' . strlen($csv_content));
        
        echo $csv_content;
        exit;

    } catch (Exception $e) {
        // Garante que a memória seja liberada mesmo em caso de erro
        if (isset($spreadsheet)) {
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
        }
        if (isset($reader)) {
            unset($reader);
        }
        
        http_response_code(500);
        echo 'Erro ao processar o arquivo: ' . $e->getMessage();
    }

} else {
    http_response_code(400);
    echo 'Nenhum arquivo enviado ou erro no upload.';
}

?>