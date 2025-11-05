Ferramenta WEB interna para automatizar completamente o processo diário de tratamento dos mapas de doação que recebemos em formato Excel.

# - O Problema que Resolve - #
Até agora, o processo era 100% manual e repetitivo. Exigia que um operador:

1) Abrisse o Excel na aba "MAP MACR".
2) Conferisse e corrigisse manualmente a coluna "HOSPITAL" com base no código "DOACAO HEMOVIDA".
3) Procurasse e removesse manualmente todas as linhas que contivessem termos como "NULO", "NULA", "CANCELADO", "INACESSIBILIDADE DE VEIA", etc.
4) Verificasse visualmente se as datas de nascimento estavam corretas.
5) Copiasse os dados até a coluna "P.A.I".
6) Salvasse como CSV e, por fim, renomeasse para .txt.

Esse fluxo era demorado e altamente suscetível a erros humanos.

# - A Solução - #
A nova ferramenta é uma página web simples que substitui todo esse trabalho.

Como funciona para o usuário:

1) O usuário acessa a página.
2) Arrasta (ou seleciona) o arquivo Excel bruto (.xls, .xlsx, etc.) para a área indicada.
3) O download do arquivo importacao.txt (já limpo, formatado e pronto para importação) é iniciado automaticamente.

# - O que o Sistema Faz (No Backend) - #
Ao receber o arquivo, o script da aplicação executa todas as regras de negócio automaticamente:

1) Validação de Datas (Crítica): Ele verifica a coluna "DT. NASC.". Se encontrar qualquer data que não seja um formato válido (ex: um texto ou formato incorreto), ele interrompe o processo e exibe um erro detalhado para o usuário.
2) Limpeza de Linhas: Ele filtra e remove automaticamente todas as linhas que contenham os termos de descarte ("NULO", "CANCELADO", etc.).
3) Aviso de Filtro: Se linhas forem removidas, o sistema gera o .txt normalmente, mas exibe um aviso amarelo informando o usuário sobre quantas e quais linhas foram filtradas.
4) Correção de Dados: Ele sobrescreve o campo "HOSPITAL" usando o 2º e 3º dígito do código "DOACAO HEMOVIDA".
5) Tratamento de Erros: É analisado se o código da "DOACAO HEMOVIDA" e o código do "DOAÇAO NAT" estão seguindo o padrão de tamanho e se não estão sendo repetidos.
6) Formatação de Saída: Ele gera o .txt final usando ponto e vírgula (;) como delimitador e garante a formatação correta (ex: sem aspas no nome do doador).

# - Detalhes Técnicos - #

1) Frontend: HTML5, CSS3 e JavaScript (Vanilla JS) para a interface de upload e feedback do usuário.
2) Backend: PHP.
3) Dependência Principal: phpoffice/phpspreadsheet (gerenciado via Composer) para ler e processar os arquivos Excel.

O objetivo é eliminar o trabalho manual, reduzir erros de importação a zero e garantir um processo padronizado e rápido.
