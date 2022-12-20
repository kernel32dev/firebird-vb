
# firebird-vb
Implementação da API Firebird em VB6

Feita para permitir qualquer aplicação em VB6 á acessar bancos de dados firebird sem a necessidade de instalar um ADO ou dependência COM

Essa implementação exige apenas que a `fbclient.dll` esteja presente ao lado do executável sem a necessiade de nenhuma instalação ou privilégios

## Instalação
1. Copie os arquivos `Firebird.cls`, `FirebirdDB.cls` e `FirebirdMod.bas` e os inclua no seu projeto
2. Baixe a ultima versão da `fbclient.dll` no [site official do Firebird](https://firebirdsql.org/en/server-packages/) e coloque ela ao lado do seu executável/projeto

## Suporte
Desenvolvido para a versão 2.5.9 do Firebird

Versões mais recentes também funcionam, já que o Firebird mantem a ABI consistente

