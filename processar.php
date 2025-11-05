<?php
// 1. Incluir o autoloader do Composer
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

// Definições de Colunas
const COLUNA_FINAL_INDEX = 12; // L é a 12ª
const COLUNA_DOADOR_INDEX = 3;  // Doador é a 4ª (índice 3)
const COLUNA_DTNASC_INDEX = 4; // DT. NASC. é a 5ª (índice 4)
const COLUNA_TUBO_NOME = 'DOAÇAO TUBO NAT'; // **NOVO**

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

// Contadores para os avisos
$linhas_removidas_count = 0;
$termos_encontrados = [];

// **MUDANÇA: Arrays para TODOS os erros de validação**
$date_validation_errors = [];
$doacao_validation_errors = []; // Para Req 1 e 2
$codigos_vistos = []; // Para Req 2 (duplicidade)

$tubo_validation_errors = []; // **NOVO (Req 3)**
$tubos_vistos = []; // **NOVO (Req 3)**


// 2. Verificar se um arquivo foi enviado
if (isset($_FILES['arquivo_excel']) && $_FILES['arquivo_excel']['error'] == UPLOAD_ERR_OK) {

    $inputFileName = $_FILES['arquivo_excel']['tmp_name'];

    try {
        // 3. Carregar o arquivo Excel
        $spreadsheet = IOFactory::load($inputFileName);
        $sheet = $spreadsheet->getSheetByName('MAP MACR');

        if (!$sheet) {
            throw new Exception("Erro: A planilha 'MAP MACR' não foi encontrada.");
        }

        // 4. Encontrar a linha do cabeçalho
        $headerRow = null; 
        $startRow = null;  
        $coluna_tubo_index = -1; // **NOVO**
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
            throw new Exception("Não foi possível encontrar a linha de cabeçalho 'DOACAO HEMOVIDA'.");
        }

        // 5. Pegar os cabeçalhos
        for ($col = 1; $col <= COLUNA_FINAL_INDEX; $col++) {
            // **CORREÇÃO APLICADA AQUI**
            $valor_cabecalho = trim((string) $sheet->getCell([$col, $headerRow])->getValue());
            $cabecalhos_array[] = $valor_cabecalho;
            
            // **NOVO: Encontra o índice da coluna do TUBO**
            if (strtoupper($valor_cabecalho) === COLUNA_TUBO_NOME) {
                $coluna_tubo_index = $col - 1; // -1 porque array é 0-indexed
            }
        }
        
        // **NOVO: Verifica se encontrou a coluna do Tubo**
        if ($coluna_tubo_index === -1) {
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

            // Pula linha se o código da doação estiver vazio
            if (empty($doacao_code)) {
                continue; 
            }

            // Limpa o ".0" que o Excel às vezes adiciona a números
            if (str_ends_with($doacao_code, ".0")) {
                $doacao_code = substr($doacao_code, 0, -2);
            }

            $erro_formato = false;
            if (!ctype_digit($doacao_code)) { // Verifica se é composto *apenas* de números
                $erro_formato = "não é composto apenas por números";
            } else if (strlen($doacao_code) != 11) {
                $erro_formato = "não possui 11 dígitos (encontrado: " . strlen($doacao_code) . ")";
            } else if (substr($doacao_code, 0, 1) != '6') {
                $erro_formato = "não começa com o dígito '6'";
            }
            
            if ($erro_formato) {
                $doacao_validation_errors[] = "Linha $row (Excel): Código '{$doacao_code_raw}' é inválido ($erro_formato).";
                continue; // Pula esta linha
            }
            
            // (Req 2): Validação de Duplicidade
            if (isset($codigos_vistos[$doacao_code])) {
                $linha_anterior = $codigos_vistos[$doacao_code];
                $doacao_validation_errors[] = "Linha $row (Excel): Código '$doacao_code' está duplicado (visto na linha $linha_anterior).";
                continue; // Pula esta linha
            }
            $codigos_vistos[$doacao_code] = $row; // Armazena a linha atual

            // === INÍCIO DA NOVA VALIDAÇÃO (DOAÇAO TUBO NAT) ===
            
            // **NOVO (Req 3): Validação do CÓDIGO TUBO NAT**
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
                continue; // Pula esta linha
            }

            // **NOVO (Req 3): Validação de Duplicidade do Tubo**
            if (isset($tubos_vistos[$tubo_code])) {
                $linha_anterior = $tubos_vistos[$tubo_code];
                $tubo_validation_errors[] = "Linha $row (Excel): Código Tubo '$tubo_code' está duplicado (visto na linha $linha_anterior).";
                continue; // Pula esta linha
            }
            $tubos_vistos[$tubo_code] = $row; // Armazena a linha atual
            
            // === FIM DA NOVA VALIDAÇÃO ===


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
                // Somente adiciona aos dados válidos se passou em TUDO
                $dados_validos[] = $linha_temporaria;
            }
        } // Fim do loop das linhas

        // 7. **MUDANÇA: VERIFICAR ERROS CRÍTICOS (Doação, Tubo e Data)**
        // Combina todos os erros críticos encontrados
        $all_critical_errors = array_merge($doacao_validation_errors, $tubo_validation_errors, $date_validation_errors);

        if (count($all_critical_errors) > 0) {
            // Pega os 10 primeiros erros para mostrar na mensagem
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

        // Escrever cabeçalho (sem aspas, com ;)
        fwrite($handle, implode($delimiter, $cabecalhos_array) . "\r\n"); 

        // Escrever dados manualmente
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

        // Envia o aviso de linhas "NULAS" removidas (se houver)
        if ($linhas_removidas_count > 0) {
            $mensagem_aviso = "Processamento concluído. ";
            $mensagem_aviso .= "$linhas_removidas_count linha(s) foram removidas por conterem os termos: " . implode(', ', $termos_encontrados) . ".";
            
            header("X-Warning-Message: " . rawurlencode($mensagem_aviso));
            header("Access-Control-Expose-Headers: X-Warning-Message");
        }

        header('Content-Type: text/plain'); 
        header('Content-Disposition: attachment; filename="importacao.txt"');
        header('Content-Length: ' . strlen($csv_content));
        
        echo $csv_content;
        exit;

    } catch (Exception $e) {
        // Pega erros (incluindo os novos erros de data e doação)
        http_response_code(500);
        echo 'Erro ao processar o arquivo: ' . $e->getMessage();
    }

} else {
    // Tratar erro de upload
    http_response_code(400);
    echo 'Nenhum arquivo enviado ou erro no upload.';
}

?>