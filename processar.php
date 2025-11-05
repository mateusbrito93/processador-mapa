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

// Lista de termos para filtrar
$termos_invalidos = [
    'NULO',
    'NULA',
    'CANCELADA',
    'CANCELADO',
    'INACESSIBILIDADE DE VEIA',
    'DOAÇÃO INVÁLIDA',
    'APENAS AMOSTRA'
];

// Arrays de dados
$dados_validos = [];
$cabecalhos_array = [];

// Contadores para os avisos
$linhas_removidas_count = 0;
$termos_encontrados = [];

// Array para erros de data
$date_errors = [];

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
            $cabecalhos_array[] = trim((string) $sheet->getCell([$col, $headerRow])->getValue());
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
            
            // Se a linha for inválida, conta e pula
            if ($linha_contem_invalido) {
                $linhas_removidas_count++;
                if (!in_array($termo_achado, $termos_encontrados)) {
                    $termos_encontrados[] = $termo_achado;
                }
                continue; // Pula para a próxima linha do Excel
            }
            
            // Se a linha for VÁLIDA, processa e transforma
            $doacao_code = (string)(int) $linha_temporaria[0];
            
            if (empty($doacao_code)) {
                continue; // Pula linha se ID da doação for vazio
            }

            // REGRA DO HOSPITAL: Sobrescreve a coluna I (índice 8)
            $codigo_hospital_novo = substr($doacao_code, 1, 2); 
            $linha_temporaria[8] = $codigo_hospital_novo;
            
            // VALIDAÇÃO 2: Formato da Data de Nascimento
            $excel_date_serial = $linha_temporaria[COLUNA_DTNASC_INDEX];
            $linha_com_erro_de_data = false;

            if (is_numeric($excel_date_serial) && $excel_date_serial > 1) {
                try {
                    // Tenta converter o número serial
                    $datetime_obj = Date::excelToDateTimeObject($excel_date_serial);
                    $linha_temporaria[COLUNA_DTNASC_INDEX] = $datetime_obj->format('d/m/Y');
                } catch (Exception $dateEx) {
                    // Pega erro se o número for um serial inválido
                    $linha_com_erro_de_data = true; 
                }
            } else if (!empty(trim((string)$excel_date_serial))) {
                // Se NÃO for numérico E NÃO estiver vazio (ex: "01/01/2000" em texto), é um erro.
                $linha_com_erro_de_data = true;
            }
            // (Se estiver vazio, a linha é processada, assumindo que data é opcional)

            // Se encontrou erro de data, armazena e NÃO adiciona a linha
            if ($linha_com_erro_de_data) {
                // $row é o número da linha no Excel
                $date_errors[] = "Linha $row (Doação: $doacao_code) tem data inválida: '$excel_date_serial'";
            } else {
                // Somente adiciona aos dados válidos se passou em TUDO
                $dados_validos[] = $linha_temporaria;
            }
        } // Fim do loop das linhas

        // 7. VERIFICAR ERROS DE DATA
        // Se houver qualquer erro de data, interrompe o script ANTES de gerar o .txt
        if (count($date_errors) > 0) {
            // Pega os 5 primeiros erros para mostrar na mensagem
            $erros_para_mostrar = array_slice($date_errors, 0, 5);
            $mensagem_erro = "Processamento interrompido. " . count($date_errors) . " erro(s) de data encontrados (a célula deve conter uma data válida do Excel, não um texto).";
            $mensagem_erro .= " Exemplos: " . implode('; ', $erros_para_mostrar);
            
            // Lança uma exceção, que será pega pelo `catch` e enviada ao frontend
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
        // Pega erros (incluindo o novo erro de data)
        http_response_code(500);
        echo 'Erro ao processar o arquivo: ' . $e->getMessage();
    }

} else {
    // Tratar erro de upload
    http_response_code(400);
    echo 'Nenhum arquivo enviado ou erro no upload.';
}

?>