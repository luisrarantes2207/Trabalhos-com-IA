# Organiza os documentos da igreja em pastas numeradas para envio ao Pastor Isac.
# Uso (PowerShell): clique com o botão direito no arquivo > "Executar com o PowerShell",
# ou abra o PowerShell e rode:  powershell -ExecutionPolicy Bypass -File .\organizar_envio_pastor.ps1
# O script SOMENTE COPIA arquivos: nada é apagado, movido ou alterado na pasta original.
param(
    [string]$Origem = (Join-Path $env:USERPROFILE "Downloads\Documentos - Constituição Igreja IBP")
)

$Destino = Join-Path $Origem "Envio - Pastor Isac"
$Mapa = @(
    @{ Padrao = "Manual_dos_Documentos"; Pasta = "00_Leia_Primeiro"; Nome = "01_Manual_dos_Documentos.docx" }
    @{ Padrao = "Nota_Tecnica"; Pasta = "00_Leia_Primeiro"; Nome = "02_Nota_Tecnica.docx" }
    @{ Padrao = "Estatuto_Social"; Pasta = "01_Constituicao_e_Registro"; Nome = "01_Estatuto_Social_e_Declaracao_de_Fe.docx" }
    @{ Padrao = "Ata_Assembleia_Fundacao"; Pasta = "01_Constituicao_e_Registro"; Nome = "02_Ata_da_Assembleia_de_Fundacao.docx" }
    @{ Padrao = "Edital_Convocacao_Fundacao"; Pasta = "01_Constituicao_e_Registro"; Nome = "03_Edital_de_Convocacao_da_Fundacao.docx" }
    @{ Padrao = "Requerimento_Registro_RCPJ"; Pasta = "01_Constituicao_e_Registro"; Nome = "04_Requerimento_de_Registro_no_Cartorio.docx" }
    @{ Padrao = "Regimento_Interno"; Pasta = "02_Normas_Internas"; Nome = "01_Regimento_Interno.docx" }
    @{ Padrao = "Codigo_de_Conduta"; Pasta = "02_Normas_Internas"; Nome = "02_Codigo_de_Conduta.docx" }
    @{ Padrao = "Politica_Protecao_Menores_Vulneraveis"; Pasta = "02_Normas_Internas"; Nome = "03_Politica_de_Protecao_de_Criancas_e_Vulneraveis.docx" }
    @{ Padrao = "Politica_Privacidade_Protecao_Dados"; Pasta = "02_Normas_Internas"; Nome = "04_Politica_de_Privacidade_e_Protecao_de_Dados.docx" }
    @{ Padrao = "Ata_Reuniao_Diretoria_Executiva"; Pasta = "03_Assembleias_e_Conselhos"; Nome = "01_Ata_de_Reuniao_da_Diretoria.docx" }
    @{ Padrao = "Ata_Assembleia_Geral_Ordinaria"; Pasta = "03_Assembleias_e_Conselhos"; Nome = "02_Assembleia_Geral_Ordinaria.docx" }
    @{ Padrao = "Ata_Assembleia_Geral_Extraordinaria"; Pasta = "03_Assembleias_e_Conselhos"; Nome = "03_Assembleia_Geral_Extraordinaria.docx" }
    @{ Padrao = "Ata_Parecer_Conselho_Fiscal"; Pasta = "03_Assembleias_e_Conselhos"; Nome = "04_Conselho_Fiscal_Atas_e_Pareceres.docx" }
    @{ Padrao = "Ata_Conselho_Ministerial"; Pasta = "03_Assembleias_e_Conselhos"; Nome = "05_Conselho_Ministerial_Atas_e_Disciplina.docx" }
    @{ Padrao = "Relatorio_Anual_Atividades"; Pasta = "03_Assembleias_e_Conselhos"; Nome = "06_Relatorio_Anual_de_Atividades.docx" }
    @{ Padrao = "Ficha_Cadastro_Membro"; Pasta = "04_Membros"; Nome = "01_Ficha_de_Cadastro_de_Membro.docx" }
    @{ Padrao = "Termo_Admissao_Membro_Menor"; Pasta = "04_Membros"; Nome = "02_Admissao_de_Membro_Menor.docx" }
    @{ Padrao = "Carta_Transferencia_Membro"; Pasta = "04_Membros"; Nome = "03_Carta_de_Transferencia_e_Membresia.docx" }
    @{ Padrao = "Termo_Desligamento_Membro"; Pasta = "04_Membros"; Nome = "04_Desligamento_de_Membro.docx" }
    @{ Padrao = "Termo_Readmissao_Membro"; Pasta = "04_Membros"; Nome = "05_Readmissao_de_Membro.docx" }
    @{ Padrao = "Certificado_Batismo"; Pasta = "05_Atos_Religiosos"; Nome = "01_Batismo_Certificado_e_Registro.docx" }
    @{ Padrao = "Certificado_Apresentacao_Criancas"; Pasta = "05_Atos_Religiosos"; Nome = "02_Apresentacao_de_Criancas.docx" }
    @{ Padrao = "Certificado_Casamento_Religioso"; Pasta = "05_Atos_Religiosos"; Nome = "03_Casamento_Religioso.docx" }
    @{ Padrao = "Certificado_Ordenacao_Ministro"; Pasta = "05_Atos_Religiosos"; Nome = "04_Ordenacao_de_Ministro.docx" }
    @{ Padrao = "Recibo_Dizimos_Ofertas"; Pasta = "06_Financas_e_Patrimonio"; Nome = "01_Recibo_de_Dizimos_e_Ofertas.docx" }
    @{ Padrao = "Termo_Doacao_Bens"; Pasta = "06_Financas_e_Patrimonio"; Nome = "02_Doacao_de_Bens_Moveis.docx" }
    @{ Padrao = "Termo_Doacao_Imovel"; Pasta = "06_Financas_e_Patrimonio"; Nome = "03_Doacao_de_Imovel.docx" }
    @{ Padrao = "Contrato_Comodato_Templo"; Pasta = "06_Financas_e_Patrimonio"; Nome = "04_Contrato_de_Comodato_do_Templo.docx" }
    @{ Padrao = "Contrato_Locacao_Templo"; Pasta = "06_Financas_e_Patrimonio"; Nome = "05_Contrato_de_Locacao_do_Templo.docx" }
    @{ Padrao = "Contrato_Trabalho_Empregado"; Pasta = "07_Pessoas_e_Contratacoes"; Nome = "01_Contrato_de_Trabalho.docx" }
    @{ Padrao = "Contrato_Prestacao_Servicos"; Pasta = "07_Pessoas_e_Contratacoes"; Nome = "02_Contrato_de_Prestacao_de_Servicos.docx" }
    @{ Padrao = "RPA_Recibo_Pagamento_Autonomo"; Pasta = "07_Pessoas_e_Contratacoes"; Nome = "03_RPA_Recibo_de_Pagamento_a_Autonomo.docx" }
    @{ Padrao = "Termo_Adesao_Servico_Voluntario"; Pasta = "07_Pessoas_e_Contratacoes"; Nome = "04_Termo_de_Adesao_ao_Servico_Voluntario.docx" }
)

if (-not (Test-Path -LiteralPath $Origem)) {
    Write-Host "Pasta de origem não encontrada: $Origem" -ForegroundColor Red
    Write-Host "Rode novamente informando a pasta: .\organizar_envio_pastor.ps1 -Origem 'C:\caminho\da\pasta'"
    exit 1
}

$arquivos = Get-ChildItem -LiteralPath $Origem -File -Filter *.docx
$usados = @{}
$faltando = @()
foreach ($item in $Mapa) {
    # aceita variações como "Nome (1).docx"; se houver mais de uma, usa a mais recente
    $candidatos = $arquivos | Where-Object { $_.BaseName -like "*$($item.Padrao)*" } | Sort-Object LastWriteTime -Descending
    if ($candidatos) {
        $arq = $candidatos[0]
        $pasta = Join-Path $Destino $item.Pasta
        New-Item -ItemType Directory -Force -Path $pasta | Out-Null
        Copy-Item -LiteralPath $arq.FullName -Destination (Join-Path $pasta $item.Nome) -Force
        $candidatos | ForEach-Object { $usados[$_.FullName] = $true }
        Write-Host "OK   $($item.Pasta)\$($item.Nome)   <=  $($arq.Name)"
    } else {
        $faltando += "$($item.Pasta)\$($item.Nome)  (procure um arquivo com '$($item.Padrao)' no nome)"
    }
}

Write-Host ""
if ($faltando.Count -gt 0) {
    Write-Host "Documentos NÃO encontrados na pasta de origem (baixe-os e rode de novo):" -ForegroundColor Yellow
    $faltando | ForEach-Object { Write-Host "  - $_" }
}
$sobras = $arquivos | Where-Object { -not $usados.ContainsKey($_.FullName) }
if ($sobras) {
    Write-Host ""
    Write-Host "Arquivos da pasta que NÃO fazem parte do envio (confira se devem ir ao Pastor):" -ForegroundColor Cyan
    $sobras | ForEach-Object { Write-Host "  - $($_.Name)" }
}
Write-Host ""
Write-Host "Concluído. Pasta de envio: $Destino" -ForegroundColor Green
