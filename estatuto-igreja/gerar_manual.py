"""Gera o Manual dos Documentos, a Nota Técnica em .docx e a pasta de envio ao Pastor.

Uso: python3 gerar_manual.py   (execute antes o gerar_docx.py)

O manual combina uma descrição escrita de cada documento (DOCUMENTOS, abaixo) com
informações extraídas automaticamente do texto-fonte: estrutura (títulos, cláusulas
e modelos) e campos a preencher (tudo o que está entre [COLCHETES]).
"""
import re
import shutil
from collections import Counter
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from gerar_docx import DOCS, build, refs

base = Path(__file__).parent
ENVIO = base.parent / "envio-pastor-isac"

# Pastas numeradas de envio: (pasta, descrição, [chaves dos documentos na ordem])
PASTAS = [
    ("01_Constituicao_e_Registro", "Documentos para fundar a igreja e registrá-la em cartório.",
     ["estatuto", "ata_fundacao", "edital_convocacao", "requerimento_registro"]),
    ("02_Normas_Internas", "Normas que regem o funcionamento, a conduta, a proteção de menores e a privacidade.",
     ["regimento_interno", "codigo_conduta", "politica_protecao", "politica_privacidade"]),
    ("03_Assembleias_e_Conselhos", "Modelos de atas dos órgãos da igreja e relatório anual.",
     ["ata_diretoria", "ata_ago", "ata_age", "ata_conselho_fiscal", "ata_conselho_ministerial", "relatorio_anual"]),
    ("04_Membros", "Entrada, saída e movimentação de membros.",
     ["ficha_membro", "termo_membro_menor", "carta_transferencia", "termo_desligamento", "termo_readmissao"]),
    ("05_Atos_Religiosos", "Certificados e livros de registro dos atos religiosos.",
     ["certificado_batismo", "certificado_apresentacao", "certificado_casamento", "certificado_ordenacao"]),
    ("06_Financas_e_Patrimonio", "Recibos, doações e contratos do imóvel do templo.",
     ["recibo_dizimos", "termo_doacao", "termo_doacao_imovel", "contrato_comodato", "contrato_locacao"]),
    ("07_Pessoas_e_Contratacoes", "Empregados, prestadores de serviço, autônomos e voluntários.",
     ["contrato_trabalho", "contrato_servicos", "rpa", "termo_voluntariado"]),
]

# Nomes curtos dos arquivos na pasta de envio
NOME_ENVIO = {
    "estatuto": "Estatuto_Social_e_Declaracao_de_Fe",
    "ata_fundacao": "Ata_da_Assembleia_de_Fundacao",
    "edital_convocacao": "Edital_de_Convocacao_da_Fundacao",
    "requerimento_registro": "Requerimento_de_Registro_no_Cartorio",
    "regimento_interno": "Regimento_Interno",
    "codigo_conduta": "Codigo_de_Conduta",
    "politica_protecao": "Politica_de_Protecao_de_Criancas_e_Vulneraveis",
    "politica_privacidade": "Politica_de_Privacidade_e_Protecao_de_Dados",
    "ata_diretoria": "Ata_de_Reuniao_da_Diretoria",
    "ata_ago": "Assembleia_Geral_Ordinaria",
    "ata_age": "Assembleia_Geral_Extraordinaria",
    "ata_conselho_fiscal": "Conselho_Fiscal_Atas_e_Pareceres",
    "ata_conselho_ministerial": "Conselho_Ministerial_Atas_e_Disciplina",
    "relatorio_anual": "Relatorio_Anual_de_Atividades",
    "ficha_membro": "Ficha_de_Cadastro_de_Membro",
    "termo_membro_menor": "Admissao_de_Membro_Menor",
    "carta_transferencia": "Carta_de_Transferencia_e_Membresia",
    "termo_desligamento": "Desligamento_de_Membro",
    "termo_readmissao": "Readmissao_de_Membro",
    "certificado_batismo": "Batismo_Certificado_e_Registro",
    "certificado_apresentacao": "Apresentacao_de_Criancas",
    "certificado_casamento": "Casamento_Religioso",
    "certificado_ordenacao": "Ordenacao_de_Ministro",
    "recibo_dizimos": "Recibo_de_Dizimos_e_Ofertas",
    "termo_doacao": "Doacao_de_Bens_Moveis",
    "termo_doacao_imovel": "Doacao_de_Imovel",
    "contrato_comodato": "Contrato_de_Comodato_do_Templo",
    "contrato_locacao": "Contrato_de_Locacao_do_Templo",
    "contrato_trabalho": "Contrato_de_Trabalho",
    "contrato_servicos": "Contrato_de_Prestacao_de_Servicos",
    "rpa": "RPA_Recibo_de_Pagamento_a_Autonomo",
    "termo_voluntariado": "Termo_de_Adesao_ao_Servico_Voluntario",
}

# Descrição de cada documento: finalidade, base legal, quem prepara/aprova/assina,
# regras de preenchimento, guarda e prazos, e cuidados.
DOCUMENTOS = {
    "estatuto": dict(
        titulo="Estatuto Social e Declaração de Fé",
        finalidade="Ato constitutivo da igreja. Define natureza jurídica, finalidades, órgãos, direitos e deveres dos membros, patrimônio, controles financeiros e dissolução, e reúne as cláusulas que sustentam a imunidade tributária. A Declaração de Fé (Anexo I) fixa os fundamentos doutrinários.",
        base="Código Civil, arts. 44, IV e § 1º, 46, 48-A e 54 a 61; CF, art. 150, VI, \"b\" (EC 132/2023) e art. 156, § 1º-A; CTN, arts. 9º e 14; LC 214/2025, art. 9º; LC 227/2026, art. 149; Lei 6.015/1973, arts. 114 e 120; Lei 8.212/1991, art. 22, §§ 13 e 14.",
        quem="Aprovado pela Assembleia Geral de Fundação (ou de reforma), com visto de advogado inscrito na OAB, e registrado no Registro Civil de Pessoas Jurídicas (RCPJ). Alterações posteriores exigem Assembleia Extraordinária com quórum qualificado.",
        regras=[
            "Preencha nome, sigla, sede e as opções doutrinárias da Declaração de Fé (seções IV e VIII).",
            "Decida os parâmetros entre colchetes: carência para votar (sugestão: 6 meses), tempo de membresia para cargos (2 anos), mandatos e reconduções, limite de endividamento, valor que obriga auditoria, mínimo de membros que impede a dissolução e a entidade que recebe o patrimônio na dissolução.",
            "Escolha se o Pastor Titular será presidente nato ou eleito para a Diretoria, e se terá mandato fixo.",
            "Não altere as cláusulas protegidas do art. [[qref]], § 3º (natureza religiosa, não distribuição, aplicação dos recursos e destino do patrimônio): elas sustentam a imunidade.",
            "Rubrique todas as folhas (presidente e secretário da Assembleia) e colha o visto do advogado na última página.",
        ],
        guarda="Via registrada guardada permanentemente na Secretaria, com cópia digital. Toda alteração deve ser averbada no RCPJ e consolidada em nova versão.",
        cuidados="O art. [[trans]], § 1º, dispensa os prazos de membresia só na fundação. Toda remissão a artigo é gerada automaticamente: se editar o Word diretamente e inserir ou remover artigos, as remissões de todos os documentos ficarão erradas; prefira editar o .txt e gerar de novo.",
    ),
    "ata_fundacao": dict(
        titulo="Ata da Assembleia Geral de Fundação",
        finalidade="Registra a fundação da igreja, a aprovação do Estatuto e da Declaração de Fé, o reconhecimento dos fundadores como membros, as eleições e a posse, a sede, o patrimônio inicial e as autorizações para registro, CNPJ e imunidades. Traz a lista de presença (Anexo A) e os termos de posse (Anexo B).",
        base="Código Civil, arts. 45 e 46; Lei 6.015/1973, arts. 114, 120 e 121; Estatuto, art. [[trans]].",
        quem="Lavrada pelo secretário da Assembleia; assinada pelo presidente e pelo secretário da Assembleia, pelos eleitos e visada por advogado. Os fundadores assinam a lista de presença.",
        regras=[
            "A qualificação completa de cada eleito (nome, nacionalidade, estado civil, profissão, RG, CPF e endereço) é obrigatória para o registro.",
            "Os resultados das votações devem ser coerentes com o número de presentes.",
            "Apague as alternativas que não se aplicam e não deixe espaços em branco nem rasuras.",
            "Fundadores devem ser civilmente capazes (maiores de 18 anos).",
            "Um Termo de Posse (Anexo B) para cada eleito, inclusive suplentes do Conselho Fiscal.",
        ],
        guarda="Registrada no RCPJ com o Estatuto; original arquivado permanentemente.",
        cuidados="Principal motivo de devolução pelo cartório: qualificação incompleta dos eleitos. Confirme com o cartório local a exigência de firma reconhecida e o número de vias.",
    ),
    "edital_convocacao": dict(
        titulo="Edital de Convocação da Assembleia de Fundação",
        finalidade="Convoca os interessados para a Assembleia de Fundação, informa a ordem do dia e as regras de participação. Inclui a certidão de afixação e divulgação.",
        base="CF, art. 5º, VI e XVII; Código Civil, art. 44, IV; Lei 13.709/2018 (LGPD).",
        quem="Assinado pela Comissão Organizadora (a igreja ainda não existe antes do registro).",
        regras=[
            "Publique com pelo menos 15 dias de antecedência.",
            "A ordem do dia deve ser idêntica à da ata de fundação.",
            "Informe onde a minuta do Estatuto pode ser consultada e o prazo para sugestões.",
            "Preencha a certidão de afixação e guarde capturas de tela da divulgação digital.",
        ],
        guarda="Arquivado com a ata de fundação; cópia apresentada ao cartório.",
        cuidados="A ata faz referência ao edital; sem a prova da convocação, a fundação fica vulnerável a questionamentos.",
    ),
    "requerimento_registro": dict(
        titulo="Requerimento de Registro no Cartório (RCPJ)",
        finalidade="Pede ao Registro Civil de Pessoas Jurídicas o registro do Estatuto e da ata de fundação e indica onde estão as informações exigidas por lei.",
        base="Código Civil, arts. 44 e 45; Lei 6.015/1973, arts. 114, 119, 120 e 121.",
        quem="Assinado pelo Presidente eleito na fundação.",
        regras=[
            "Informe e-mail e telefone para o cartório comunicar exigências.",
            "Indique o apresentante (o próprio requerente ou outra pessoa com endereço).",
            "Confira a lista de documentos com o cartório da comarca antes de protocolar e faça pesquisa de nome.",
        ],
        guarda="Cópia protocolada arquivada com os atos constitutivos.",
        cuidados="A Lei 6.015, art. 121, exige hoje uma via (papel ou eletrônica), mas alguns cartórios ainda pedem mais.",
    ),
    "regimento_interno": dict(
        titulo="Regimento Interno",
        finalidade="Regulamenta o que o Estatuto delega: admissão e disciplina de membros, assembleias e eleições, funcionamento dos órgãos, alçadas financeiras, ministérios, congregações, integridade, proteção de vulneráveis, LGPD, celebrações e mediação.",
        base="Estatuto, arts. [[reg]] e [[agc]], X; ECA, arts. 13, 59-A e 70-B; Lei 9.504/1997, art. 37, § 4º; LGPD.",
        quem="Proposto pela Diretoria (ouvido o Conselho Ministerial) e aprovado pela Assembleia Geral por maioria simples, em até 180 dias da fundação. Não precisa de registro em cartório.",
        regras=[
            "Defina os valores das alçadas de despesa (art. <<alc>>), do fundo fixo de caixa (art. <<caixa>>) e de alienação de bens móveis.",
            "Defina prazos de medidas disciplinares, número de presbíteros/oficiais e a lista de ministérios.",
            "Revise a cada 4 anos.",
        ],
        guarda="Versão aprovada e ata de aprovação arquivadas; cópia disponível aos membros.",
        cuidados="Não pode contrariar o Estatuto. A regra eleitoral cita a Lei 9.504, art. 37, § 4º (templos como bens de uso comum).",
    ),
    "codigo_conduta": dict(
        titulo="Código de Conduta",
        finalidade="Padrões de conduta para dirigentes, ministros, líderes, voluntários, empregados e prestadores: respeito, autoridade espiritual, aconselhamento, recursos, conflitos de interesse, comunicação, neutralidade política, segurança e consequências.",
        base="Estatuto, art. [[int]], I; Regimento, arts. <<cc1>> e <<fin1>>; Lei 12.846/2013.",
        quem="Aprovado pela Diretoria, ouvido o Conselho Ministerial. Cada colaborador assina o Termo de Ciência anexo.",
        regras=[
            "Defina o valor máximo de brindes aceitáveis.",
            "Decida se mantém a regra opcional de afastamento de candidatos a cargo eletivo durante a campanha.",
            "Colha o Termo de Ciência no início de cada função e a cada renovação.",
        ],
        guarda="Termos de Ciência arquivados na Secretaria enquanto durar a função e depois pelo prazo da Política de Privacidade.",
        cuidados="A proibição de prometer bênçãos em troca de contribuições reforça o caráter voluntário das receitas, importante para a imunidade.",
    ),
    "politica_protecao": dict(
        titulo="Política de Proteção de Crianças, Adolescentes e Pessoas Vulneráveis",
        finalidade="Prevenção, identificação e resposta a violências nas atividades da igreja: governança, seleção de colaboradores, regras de atividades, viagens, imagem, resposta a revelações e participação de pessoas com histórico de violência. Anexos: termo do colaborador, ficha da criança, autorização de viagem e registro de incidente.",
        base="CF, art. 227; ECA, arts. 2º, 13, 18-A, 59-A, 70-B e 83; Lei 13.431/2017, arts. 4º e 13; LGPD, art. 14.",
        quem="Aprovada pela Diretoria, ouvido o Conselho Ministerial. Aplicada pelo(a) Coordenador(a) de Proteção.",
        regras=[
            "Designe o(a) Coordenador(a) de Proteção e suplente (de preferência sem ligação com o Pastor Titular).",
            "Ajuste a proporção de crianças por adulto ao espaço e ao número de voluntários.",
            "Certidões criminais dos colaboradores renovadas a cada 12 meses (6 meses se houver recursos públicos).",
            "Nas viagens, confira a resolução do CNJ vigente sobre autorização de viagem.",
        ],
        guarda="Registros de incidentes: 20 anos, sigilosos. Certidões: durante a função e mais 5 anos.",
        cuidados="A comunicação ao Conselho Tutelar em caso de suspeita é obrigatória e imediata, e nenhum órgão da igreja pode impedi-la ou atrasá-la.",
    ),
    "politica_privacidade": dict(
        titulo="Política de Privacidade e Proteção de Dados",
        finalidade="Aviso aos titulares (dados tratados, finalidades, bases legais, compartilhamentos, prazos de guarda e direitos) e normas internas (responsabilidades, segurança, operadores e incidentes). Anexos: consentimento do membro, formulário do titular e cláusula para fornecedores.",
        base="LGPD (Lei 13.709/2018), arts. 5º, 7º, 11, 14, 16, 18, 19, 33, 37, 38, 41, 46 e 48; Resoluções CD/ANPD nº 2/2022 e 15/2024.",
        quem="Aprovada pela Diretoria. O encarregado de dados é designado em ata da Diretoria.",
        regras=[
            "Preencha os dados e o e-mail exclusivo do encarregado.",
            "Ajuste a tabela de dados às atividades reais da igreja (retire site, app ou ação social se não existirem).",
            "Defina os prazos de guarda entre colchetes.",
            "Publique a versão vigente no site e na Secretaria, com a data de atualização.",
        ],
        guarda="Versões anteriores arquivadas. Registro de incidentes por pelo menos 5 anos.",
        cuidados="Religião é dado sensível: o simples cadastro de membro revela a crença. Nunca use legítimo interesse para dados sensíveis. Incidentes: comunicação à ANPD em 3 dias úteis (6 para pequeno porte).",
    ),
    "ata_diretoria": dict(
        titulo="Ata de Reunião da Diretoria Executiva",
        finalidade="Modelo de ata mensal da Diretoria, com presenças, relatório financeiro padronizado, deliberações, quadro de providências e textos prontos para as decisões mais comuns.",
        base="Estatuto, arts. [[de]], § 4º, [[ci]] e [[mov]]; Regimento, arts. <<dir1>> e <<alc>>.",
        quem="Lavrada pelo 1º Secretário e assinada por todos os presentes.",
        regras=[
            "Numere as atas em sequência a cada ano (ex.: 07/2027).",
            "Registre por completo decisões, valores, prazos e responsáveis, e o resultado de cada votação.",
            "Registre obrigatoriamente: despesas acima da alçada, perfis de acesso bancário, negócios com partes relacionadas, doações relevantes, designações e convocações de Assembleia.",
            "Não registre informações pastorais, de saúde ou disciplinares de membros.",
        ],
        guarda="Livro ou arquivo de atas, permanente.",
        cuidados="A ata é a prova de que os controles do Estatuto foram cumpridos perante o Conselho Fiscal e o Fisco.",
    ),
    "ata_ago": dict(
        titulo="Assembleia Geral Ordinária",
        finalidade="Calendário preparatório, edital e ata da AGO anual (contas, orçamento e eleições), com lista de presença.",
        base="Estatuto, arts. [[agc]] a [[ata]]; Regimento, arts. <<ag1>> a <<ag5>> e <<el1>> a <<el4>>.",
        quem="Convocada pelo Presidente; presidida pelo Presidente, com presidente e secretário ad hoc no item das contas; ata assinada pela mesa e pelos eleitos.",
        regras=[
            "Siga o calendário de prazos da primeira página (Conselho Fiscal 30 dias antes, documentos 15 dias antes, edital 8 ou 15 dias antes).",
            "Eleja presidente e secretário ad hoc para as contas (Estatuto, art. [[quo]], § 3º); recomende que diretores não votem as próprias contas.",
            "Registre a qualificação completa dos eleitos e autorize o registro da ata no cartório.",
            "Recurso disciplinar: só o resultado na ata geral; detalhes em ata reservada.",
        ],
        guarda="Ata registrada no RCPJ quando houver eleição; original permanente.",
        cuidados="Realizar até 30 de abril de cada ano. O parecer do Conselho Fiscal lido na AGO é a prova anual da imunidade.",
    ),
    "ata_age": dict(
        titulo="Assembleia Geral Extraordinária",
        finalidade="Tabela de quóruns por matéria, edital, ata com módulos (reforma estatutária, destituição, cargo vago, imóveis, Declaração de Fé, filiação, dissolução), lista de presença e termo de verificação de quórum.",
        base="Código Civil, arts. 59 e 60; Estatuto, arts. [[agc]] a [[ata]], [[vac]], [[dest0]] e [[alt]] a [[dest]].",
        quem="Convocada pelo Presidente, pela Diretoria, pelo Conselho Fiscal ou por 1/5 dos membros.",
        regras=[
            "Use só os módulos que correspondem à pauta.",
            "Descreva a matéria de forma específica no edital (na destituição, os fatos da justa causa) e publique com 15 dias de antecedência.",
            "Confira a presença mínima antes de votar; se não atingida, faça nova convocação com 15 dias de intervalo.",
            "Preencha o Termo de Verificação de Quórum em toda votação qualificada.",
        ],
        guarda="Ata averbada no RCPJ quando alterar o Estatuto, eleger ou dissolver; original permanente.",
        cuidados="Edital genérico, falta de defesa na destituição e quórum mal contado são as causas mais comuns de anulação judicial.",
    ),
    "ata_conselho_fiscal": dict(
        titulo="Conselho Fiscal: Atas e Pareceres",
        finalidade="Ata trimestral com roteiro de verificações, parecer anual sobre as contas com atestado da imunidade, parecer específico e comunicação de irregularidade à Diretoria.",
        base="Estatuto, arts. [[imu]], § 1º, [[dest0]], [[cf]] e [[ci]]; Regimento, art. <<cfis>>.",
        quem="Conselho Fiscal (titulares e, nas faltas, suplentes).",
        regras=[
            "Aplique o mesmo roteiro de verificações todo trimestre.",
            "Entregue o parecer anual até 30 dias antes da AGO.",
            "Conselheiro com interesse pessoal declara impedimento e é substituído.",
        ],
        guarda="Permanente, com as demonstrações do exercício.",
        cuidados="O Conselho fiscaliza, mas não administra nem trata de assuntos pastorais.",
    ),
    "ata_conselho_ministerial": dict(
        titulo="Conselho Ministerial: Atas e Procedimento Disciplinar",
        finalidade="Ata geral (admissões, desligamentos, líderes, credenciais, celebrações) e documentos reservados do procedimento disciplinar: instauração, notificação, decisão e comunicação.",
        base="Estatuto, arts. [[disc]] e [[cm]]; Regimento, arts. <<fal>> a <<pd7>> e <<ord>>; LGPD; ECA.",
        quem="Presidido pelo Pastor Titular (ou pelo ministro mais antigo se ele for o acusado).",
        regras=[
            "Separe ata geral e ata reservada.",
            "Descreva os fatos de forma objetiva, sem adjetivos, e fundamente a decisão.",
            "Garanta ao acusado acesso aos autos, 10 dias para defesa, testemunhas e acompanhante ou advogado.",
            "Informe o prazo de 15 dias para recurso à Assembleia.",
        ],
        guarda="Autos disciplinares sigilosos por 10 anos; ata geral permanente.",
        cuidados="Suspeita de violência contra vulnerável é comunicada imediatamente às autoridades, independentemente do procedimento.",
    ),
    "relatorio_anual": dict(
        titulo="Relatório Anual de Atividades",
        finalidade="Relatório da Diretoria para a AGO, organizado pelas finalidades da igreja, com números, resumo financeiro por finalidade, indicadores, conformidade e declaração sobre os requisitos da imunidade.",
        base="Estatuto, arts. [[fin]], [[imu]], [[agc]], IV, [[dec]], II, e [[cont]], § 1º.",
        quem="Elaborado pela Diretoria; assinado por Presidente, 1º Tesoureiro, 1º Secretário e Pastor Titular.",
        regras=[
            "Os números devem conferir com a contabilidade.",
            "Cada informação deve ter documento arquivado.",
            "Ninguém é identificado nas seções sensíveis; use só números.",
            "Seções sem atividade ficam com a indicação \"sem atividade no exercício\".",
        ],
        guarda="Permanente, com as contas do exercício.",
        cuidados="É a principal prova, perante o Fisco, de que a igreja se dedica às finalidades essenciais.",
    ),
    "ficha_membro": dict(
        titulo="Ficha de Cadastro de Membro",
        finalidade="Coleta os dados necessários ao Rol de Membros, dados facultativos, declaração de adesão, consentimento LGPD, autorização dos pais (menores), campos da Secretaria e atualização anual.",
        base="Estatuto, arts. [[adm]] e [[cat]]; Regimento, arts. <<adm1>> e <<rol>>; LGPD, arts. 11 e 14.",
        quem="Preenchida pelo candidato (ou pelos pais, se menor) antes da entrevista; Parte E pela Secretaria.",
        regras=[
            "Parte A é obrigatória; Parte B é facultativa e a recusa não impede a admissão.",
            "Não peça dados de saúde, renda ou opinião política.",
            "Medidas disciplinares não vão na ficha.",
            "Aplique a Parte F (atualização) todo ano, antes da AGO.",
        ],
        guarda="Armário trancado ou sistema com acesso individual; após desligamento, só nome e datas.",
        cuidados="O consentimento pode ser revogado a qualquer tempo, conservando-se apenas os registros exigidos por lei ou pelo Estatuto.",
    ),
    "termo_membro_menor": dict(
        titulo="Termo de Admissão de Membro Menor de 18 Anos",
        finalidade="Autorização dos pais, manifestação do próprio menor, comunicação ao outro genitor e ratificação aos 18 anos.",
        base="ECA, art. 16, III; Código Civil (poder familiar); LGPD, art. 14; Regimento, art. <<adm1>>, § 1º.",
        quem="Pais ou responsáveis (Parte A), o menor (Parte B), Conselho Ministerial (registro da conversa).",
        regras=[
            "Defina a idade mínima conforme a doutrina.",
            "Colha a assinatura dos dois pais sempre que possível.",
            "A conversa com o menor ocorre em local visível, com os pais presentes ou cientes.",
            "Havendo oposição de um dos pais, suspenda a admissão.",
        ],
        guarda="Com a ficha do membro.",
        cuidados="Nenhum menor é admitido contra a vontade ou sob pressão. Aos 18 anos, colha o consentimento em nome próprio.",
    ),
    "carta_transferencia": dict(
        titulo="Carta de Transferência e Documentos de Membresia",
        finalidade="Pedido e carta de transferência, comunicação de recebimento, carta de recomendação, declaração de membresia e solicitação de carta à igreja de origem.",
        base="Estatuto, arts. [[adm]], [[cat]] e [[disc]]; Regimento, arts. <<adm2>> e <<rol>>; LGPD.",
        quem="Emitida por deliberação do Conselho Ministerial; assinada pelo Pastor Titular e pelo 1º Secretário.",
        regras=[
            "Só a pedido do membro e só para membro em plena comunhão.",
            "Dados mínimos: nome, datas, modo de admissão e batismo.",
            "O membro sai do Rol quando a igreja de destino confirmar o recebimento.",
            "Membro suspenso recebe apenas Declaração de Membresia, sem motivos.",
        ],
        guarda="Cópia com a ficha do membro.",
        cuidados="Nunca informar fatos pastorais, de saúde, financeiros ou disciplinares.",
    ),
    "termo_desligamento": dict(
        titulo="Desligamento de Membro",
        finalidade="Pedido de desligamento, termo de desligamento (qualquer hipótese) com providências da Secretaria, notificação por ausência prolongada e resposta ao ex-membro.",
        base="CF, art. 5º, VI e XX; Estatuto, arts. [[dir]], VII, e [[disc]]; Regimento, art. <<aus>>.",
        quem="Membro (pedido); Secretaria e Conselho Ministerial (termo); Pastor Titular (notificação e resposta).",
        regras=[
            "Nunca condicione a saída a entrevista ou quitação de contribuições.",
            "Cumpra todas as providências do quadro do Modelo 2.",
            "Na ausência prolongada, aguarde 15 dias após a notificação.",
        ],
        guarda="Termo com a ficha; após o prazo da Política de Privacidade, só nome e datas.",
        cuidados="O motivo da saída nunca é divulgado. Não há devolução de contribuições.",
    ),
    "termo_readmissao": dict(
        titulo="Readmissão de Membro",
        finalidade="Pedido de readmissão, termo de readmissão, compromisso de reparação de dano e comunicação da decisão.",
        base="Estatuto, arts. [[adm]], [[cat]] e [[req]]; Regimento, arts. <<adm2>> a <<adm4>> e <<pd4>>.",
        quem="Conselho Ministerial decide; Presidente assina o termo de reparação.",
        regras=[
            "Use o modo conforme a saída: aclamação (a pedido ou ausência), transferência ou reconciliação (exclusão).",
            "Os prazos de carência recomeçam; exija nova ficha e novo consentimento.",
            "Reparação de dano é contabilizada como ressarcimento, não como oferta.",
        ],
        guarda="Com a ficha do membro.",
        cuidados="Casos de violência contra vulneráveis exigem acordo de participação supervisionada. Exclusão anulada em recurso não é readmissão.",
    ),
    "certificado_batismo": dict(
        titulo="Certificado e Registro de Batismo",
        finalidade="Certificado numerado, livro de registro de batismos e declaração de segunda via.",
        base="Declaração de Fé, seção VIII; Regimento, art. <<adm1>>, § 1º; LGPD.",
        quem="Ministro celebrante e Pastor Titular (certificado); Secretaria (livro).",
        regras=[
            "Numere o certificado com o livro e a folha do registro.",
            "Não registre CPF nem endereço no livro.",
            "Correções só por anotação datada, nunca por rasura.",
            "Ajuste forma, fórmula e versículo à doutrina.",
        ],
        guarda="Livro de batismos permanente, com acesso restrito.",
        cuidados="Batismo de menor exige autorização dos pais e manifestação do menor. O certificado não substitui documentos civis.",
    ),
    "certificado_apresentacao": dict(
        titulo="Apresentação de Crianças",
        finalidade="Pedido dos pais (com autorização de imagem), certificado, livro de registro e roteiro litúrgico.",
        base="LGPD, art. 14; Regimento, art. <<cel>>; Política de Proteção.",
        quem="Pais ou responsáveis (pedido); ministro celebrante e Pastor Titular (certificado).",
        regras=[
            "Só a pedido dos responsáveis, de preferência os dois.",
            "Respeite a opção de imagem marcada pelos pais na transmissão.",
            "O ato é gratuito.",
        ],
        guarda="Livro permanente, com acesso restrito.",
        cuidados="Não é batismo nem admissão como membro.",
    ),
    "certificado_casamento": dict(
        titulo="Casamento Religioso",
        finalidade="Pedido de celebração, Termo de Casamento Religioso para registro civil, requerimento ao cartório, certificado e livro de registro.",
        base="Código Civil, arts. 1.515, 1.516 e 1.532; Lei 6.015/1973, arts. 71 a 75; Regimento, art. <<cel>>; Declaração de Fé, seção IX.",
        quem="Noivos (pedido); celebrante, contraentes e duas testemunhas (Termo); celebrante (requerimento).",
        regras=[
            "Com efeito civil, só celebre com a certidão de habilitação conferida e dentro da validade.",
            "Preencha o Termo sem rasuras, com a qualificação completa dos contraentes e testemunhas.",
            "Protocole o registro no cartório o quanto antes e confirme o prazo e a exigência de firma reconhecida.",
        ],
        guarda="Livro permanente; cópia do protocolo e da certidão de registro.",
        cuidados="Sem registro no prazo, o casamento não tem efeito civil. O certificado religioso não prova casamento civil.",
    ),
    "certificado_ordenacao": dict(
        titulo="Ordenação de Ministro",
        finalidade="Ata de exame e aprovação, termo de compromisso ministerial, certificado de ordenação, credencial anual e livro de registro.",
        base="Estatuto, arts. [[cm]] e [[min]]; Regimento, arts. <<ord>> e <<sus>>; Lei 8.212/1991, art. 22, §§ 13 e 14.",
        quem="Conselho Ministerial examina e aprova; o ministro assina o termo; Pastor Titular, Presidente e oficiante assinam o certificado.",
        regras=[
            "Confira todos os requisitos da tabela antes do exame.",
            "O termo de compromisso é assinado antes do culto de ordenação.",
            "Renove a credencial todo ano e recolha as vencidas ou suspensas.",
        ],
        guarda="Livro de ordenações permanente.",
        cuidados="O termo de compromisso é a principal prova contra pedido de vínculo de emprego e contra a caracterização do sustento como salário.",
    ),
    "recibo_dizimos": dict(
        titulo="Recibo de Dízimos, Ofertas e Doações",
        finalidade="Instruções à Tesouraria, recibo com canhoto, declaração anual de contribuições e requerimento de devolução por erro material.",
        base="Estatuto, arts. [[dev]], [[disc]], § 5º, [[rec]] e [[nd]]; Regimento, art. <<cont>>; Declaração de Fé, seção VI.",
        quem="Tesoureiro emite e assina; Presidente assina a declaração anual.",
        regras=[
            "Numeração sequencial em duas vias; recibo obrigatório acima do valor definido e em doações designadas.",
            "Nunca emita recibo com dados diferentes do efetivamente recebido.",
            "Devolução só por erro material, com aprovação da Diretoria e para a conta de origem.",
        ],
        guarda="10 anos, com termos de contagem e comprovantes.",
        cuidados="O recibo não serve para dedução de IR, salvo hipótese legal. Confirme com o contador antes de emitir para empresas.",
    ),
    "termo_doacao": dict(
        titulo="Doação de Bens Móveis",
        finalidade="Termo de doação de bens móveis e veículos, com roteiro para doação de imóvel.",
        base="Código Civil, arts. 538 a 564; LC 227/2026, art. 149.",
        quem="Doador (e cônjuge, se bens comuns) e Presidente da igreja.",
        regras=[
            "Descreva cada bem com estado e valor estimado.",
            "Escolha doação pura ou com encargo (com a válvula de alienação).",
            "Evite a cláusula de reversão.",
            "Para veículos, informe placa, RENAVAM e chassi e exija a transferência no órgão de trânsito.",
        ],
        guarda="Com o inventário patrimonial, enquanto o bem existir e mais 10 anos.",
        cuidados="Imóveis não se doam por este termo; use a doação de imóvel (escritura).",
    ),
    "termo_doacao_imovel": dict(
        titulo="Doação de Imóvel",
        finalidade="Minuta de escritura pública, instrumento particular (imóveis até 30 salários mínimos) e extrato de ata de aceitação.",
        base="Código Civil, arts. 108, 538 a 564, 1.245 e 1.647; LC 227/2026, art. 149.",
        quem="Tabelião (escritura), doador e cônjuge, Presidente da igreja; Diretoria aceita (e Assembleia se houver encargo relevante).",
        regras=[
            "Acima de 30 salários mínimos, sempre escritura pública.",
            "Faça as diligências (matrícula, certidões, CNIB) antes de aceitar.",
            "Peça o reconhecimento da imunidade de ITCMD antes da escritura.",
            "Registre na matrícula imediatamente.",
        ],
        guarda="Escritura e matrícula atualizada, permanentemente.",
        cuidados="A igreja só é dona após o registro. Depois, peça a imunidade do IPTU.",
    ),
    "contrato_comodato": dict(
        titulo="Contrato de Comodato do Templo",
        finalidade="Empréstimo gratuito do imóvel usado como templo, com laudo de vistoria e termo de devolução.",
        base="Código Civil, arts. 579 a 585 e 1.647; Estatuto, art. [[ci]].",
        quem="Comodante (e cônjuge) e Presidente da igreja; duas testemunhas.",
        regras=[
            "Prefira prazo determinado (sugestão: 10 anos) e aviso de 180 dias para devolução.",
            "Defina quem paga o IPTU se a imunidade for negada.",
            "Preencha o laudo de vistoria com fotos.",
            "Se o comodante for ligado a dirigente, cumpra o art. [[ci]] do Estatuto.",
        ],
        guarda="Durante o contrato e mais 10 anos.",
        cuidados="A imunidade de IPTU para imóvel em comodato é incerta; não converta em locação simbólica. O comodato não obriga quem comprar o imóvel.",
    ),
    "contrato_locacao": dict(
        titulo="Contrato de Locação do Templo",
        finalidade="Locação não residencial do imóvel usado como templo, com laudo de vistoria e termo de devolução.",
        base="Lei 8.245/1991, arts. 4º, 8º, 9º, 13, 19, 22, 23, 27 a 37, 45, 53, 56 e 63; CF, art. 156, § 1º-A; Estatuto, arts. [[agc]], VI, e [[ci]].",
        quem="Locador (e cônjuge), Presidente da igreja, fiador se houver; duas testemunhas.",
        regras=[
            "Informe o número de registro da igreja (necessário para a proteção do art. 53).",
            "Escolha uma só garantia (preferir caução ou seguro-fiança).",
            "Registre ou averbe o contrato na matrícula logo após assinar.",
            "Locação por mais de 5 anos exige autorização da Assembleia.",
        ],
        guarda="Durante o contrato e mais 10 anos.",
        cuidados="Se o locador for pessoa física, a igreja retém o IR sobre o aluguel.",
    ),
    "contrato_trabalho": dict(
        titulo="Contrato de Trabalho",
        finalidade="Contrato CLT para funções administrativas, técnicas ou operacionais, com descrição da função, banco de horas e opção de vale-transporte.",
        base="CLT, arts. 2º, 58-A, 59, 71, 73, 74, 443 a 469, 477 e 482; Lei 7.418/1985; Estatuto, art. [[ci]].",
        quem="Presidente da igreja e empregado; duas testemunhas.",
        regras=[
            "Não use para ministros.",
            "Descreva a função no Anexo I.",
            "Registre o ponto mesmo quando não obrigatório.",
            "Verifique convenção coletiva aplicável (piso e regras).",
        ],
        guarda="Documentos trabalhistas pelos prazos da legislação.",
        cuidados="Participação em cultos fora do horário é livre e não conta como jornada. É proibido descontar dízimo do salário.",
    ),
    "contrato_servicos": dict(
        titulo="Contrato de Prestação de Serviços",
        finalidade="Contrato com prestador pessoa física ou jurídica, com escopo e níveis de serviço e declarações da contratada.",
        base="Código Civil, arts. 593 a 609; Lei 6.019/1974; LGPD; Lei 12.846/2013; Estatuto, art. [[ci]].",
        quem="Presidente da igreja e contratada; duas testemunhas.",
        regras=[
            "Descreva o escopo no Anexo I.",
            "Em terceirização com mão de obra alocada, exija comprovantes mensais.",
            "Confirme as retenções com o contador.",
            "Prazo máximo de 4 anos por contrato.",
        ],
        guarda="Durante o contrato e mais 10 anos.",
        cuidados="Não trate o prestador como empregado (horário, ordens, advertências).",
    ),
    "rpa": dict(
        titulo="RPA – Recibo de Pagamento a Autônomo",
        finalidade="Recibo para pagamentos eventuais a pessoa física sem nota fiscal, com demonstrativo de retenções, declarações do prestador e quadro de controle interno.",
        base="Lei 8.212/1991 (contribuinte individual); legislação do IR e do ISS.",
        quem="Tesoureiro emite; prestador assina e preenche as declarações.",
        regras=[
            "Confira as tabelas vigentes de INSS e IR e a lei do ISS a cada pagamento.",
            "Reembolso de despesas comprovadas vai em campo separado e não é tributado.",
            "Informe no eSocial e na DCTFWeb.",
        ],
        guarda="10 anos, com as guias.",
        cuidados="\"Oferta\" fixa a pregador convidado sem comprovantes é remuneração. Pagamentos frequentes à mesma pessoa podem virar vínculo de emprego.",
    ),
    "termo_voluntariado": dict(
        titulo="Termo de Adesão ao Serviço Voluntário",
        finalidade="Formaliza o voluntariado, com registro de atividades e termo de encerramento.",
        base="Lei 9.608/1998, arts. 1º a 3º; Lei 9.610/1998; Estatuto, art. [[vol]].",
        quem="Representante da igreja e voluntário (e pais, se adolescente).",
        regras=[
            "Descreva a atividade e a disponibilidade definida pelo voluntário.",
            "Ressarcimento só de despesas autorizadas e comprovadas.",
            "Adolescentes de 14 a 17 anos só com os pais e fora de horário noturno ou perigoso.",
        ],
        guarda="Durante o serviço e mais 5 anos.",
        cuidados="Evite metas, escalas obrigatórias e ajudas de custo fixas, que caracterizam vínculo de emprego.",
    ),
}

COMUNS = ["NOME COMPLETO DA IGREJA", "SIGLA", "NÚMERO", "DATA", "NOME COMPLETO", "NOME", "ANO", "UF",
          "CIDADE", "VALOR", "ENDEREÇO COMPLETO, COM CEP", "MÊS", "NACIONALIDADE", "ESTADO CIVIL", "DIA",
          "PROFISSÃO", "ÓRGÃO/UF", "NOME DO PRESIDENTE", "HORA", "ENDEREÇO", "LOCAL", "DIA POR EXTENSO",
          " ", "", "E-MAIL", "TELEFONE", "CARGO", "VALOR POR EXTENSO", "ÓRGÃO EXPEDIDOR/UF", "NOMES",
          "ENDEREÇO COMPLETO", "QUALIFICAÇÃO", "NOME DO PASTOR TITULAR", "ASSINATURA"]


def campos(src):
    t = (base / src).read_text(encoding="utf-8")
    # remissões dentro de campos viram "(nº)", pois a numeração é resolvida só no documento
    t = re.sub(r"\[\[\w+\]\]|<<\w+>>|\{\{[\w:]+\}\}|\(\(\w+\)\)", "(nº)", t)
    ph = re.findall(r"\[(?!\[)([A-ZÀ-Ý0-9 ][^\[\]\n]*?)\](?!\])", t)
    return Counter(p.strip() for p in ph)


def estrutura(src):
    """Títulos (##), subtítulos (###) e cláusulas (CL) do documento, com referências resolvidas."""
    itens = []
    clausulas = 0
    for line in (base / src).read_text(encoding="utf-8").splitlines():
        if line.startswith("## "):
            itens.append(line[3:])
        elif line.startswith("CL["):
            clausulas += 1
            itens.append(f"Cláusula {clausulas}ª – " + re.sub(r"CL\[\w+\]\s*", "", line))
    n_art = sum(1 for l in (base / src).read_text(encoding="utf-8").splitlines() if re.match(r"R?ART\[", l))
    return itens, n_art


def escrever_manual():
    chave_arquivo = {Path(s).stem: (s, o) for s, o in DOCS}
    # chaves de DOCUMENTOS usam o nome do .txt sem extensão
    L = []
    a = L.append
    a("# MANUAL DOS DOCUMENTOS DA [NOME COMPLETO DA IGREJA] – [SIGLA]")
    a("!Como os documentos foram montados, o que contém cada um e como preenchê-los")
    a("!Versão de [DATA] – documento de apoio, não substitui a revisão por advogado")
    a("## 1. APRESENTAÇÃO")
    a("P Este manual acompanha o conjunto de 32 modelos preparado para constituir, registrar e fazer funcionar a [NOME COMPLETO DA IGREJA] como organização religiosa com imunidade tributária. Para cada documento, explica a finalidade, a base legal, quem prepara, aprova e assina, a estrutura, as regras de preenchimento, os campos a preencher, a guarda e os principais cuidados.")
    a("P O manual não substitui a revisão por advogado inscrito na OAB, que é indispensável antes do registro e antes da primeira utilização de cada contrato. Parte dos dispositivos legais citados nos documentos não pôde ser conferida no texto oficial durante a elaboração; a Nota Técnica (pasta 00) indica quais.")
    a("## 2. COMO OS DOCUMENTOS FORAM MONTADOS")
    a("P Ponto de partida: o rascunho inicial do Estatuto enviado pela igreja, reestruturado com base no Código Civil, na Constituição Federal (imunidade religiosa, com a redação da Reforma Tributária), no Código Tributário Nacional, nas Leis Complementares 214/2025 e 227/2026, na Lei de Registros Públicos, na legislação previdenciária e trabalhista, no ECA e na LGPD.")
    a("P Ordem de construção: primeiro o Estatuto e a Declaração de Fé; depois os documentos de fundação e registro; em seguida as normas internas (Regimento, Código de Conduta e Políticas), que regulamentam o que o Estatuto delega; por fim os modelos de uso diário, sempre remetendo ao Estatuto e ao Regimento para manter a coerência.")
    a("P Método técnico: cada documento foi escrito em um arquivo de texto (.txt) com marcações simples e convertido em Word (.docx) por um programa gerador. O gerador numera automaticamente os artigos e as cláusulas e resolve as remissões entre documentos (por exemplo, quando a ata de fundação cita \"art. 50 do Estatuto\"), de modo que, se o Estatuto for alterado e os artigos mudarem de número, basta gerar de novo para que todos os documentos acompanhem a nova numeração.")
    a("P Verificações feitas: todos os documentos foram gerados sem erro; toda remissão a artigo é conferida pelo gerador, que interrompe a execução se a referência não existir; e os números de artigo citados na Nota Técnica foram conferidos contra a numeração gerada. A formatação visual (quebras de página e tabelas) deve ser conferida no Word antes da impressão.")
    a("## 3. REGRAS GERAIS DE PREENCHIMENTO")
    a("I – Tudo o que está entre [COLCHETES] deve ser preenchido ou apagado. Nada entre colchetes pode permanecer no documento final.")
    a("II – Colchetes com barra (por exemplo, [SIM / NÃO] ou [PURO E SIMPLES / COM ENCARGO]) indicam alternativas: mantenha a aplicável e apague as demais.")
    a("III – Valores entre colchetes com número e extenso (por exemplo, [6 (SEIS)] meses) são sugestões: podem ser mantidos ou ajustados, sempre com número e extenso.")
    a("IV – Campos marcados com [ ] são opções para assinalar com X; não apague as opções não marcadas em formulários.")
    a("V – Trechos iniciados por [SE ...:] ou [OPCIONAL:] só se aplicam na situação indicada; apague-os quando não se aplicarem.")
    a("VI – Datas por extenso em atas e escrituras; valores em reais com número e extenso; qualificação completa de pessoas (nome, nacionalidade, estado civil, profissão, RG com órgão expedidor, CPF e endereço com CEP) sempre que o documento pedir.")
    a("VII – Não renumere artigos à mão nem insira ou remova artigos diretamente no Word: as remissões dos outros documentos ficariam erradas. Para alterar a estrutura, edite o arquivo .txt correspondente e gere os documentos de novo (ou peça a quem mantém os arquivos).")
    a("VIII – Documentos para cartório: sem rasuras, sem espaços em branco, rubricas em todas as folhas quando exigido e firmas reconhecidas conforme a exigência do cartório local.")
    a("IX – Dados pessoais: preencha apenas o que o documento pede, guarde os documentos preenchidos com acesso restrito e siga os prazos de guarda da Política de Privacidade.")
    a("X – Numeração: atas, recibos, certificados e termos são numerados em sequência por ano (por exemplo, 07/2027), e cada livro de registro tem termo de abertura e de encerramento.")
    a("## 4. CAMPOS COMUNS A VÁRIOS DOCUMENTOS")
    a("%GRID 6,10")
    a("Campo | Como preencher")
    a("[NOME COMPLETO DA IGREJA] e [SIGLA] | Exatamente como no Estatuto registrado, em todos os documentos")
    a("[DATA], [DIA], [MÊS], [ANO], [HORA] | Data do ato; em atas e escrituras, por extenso")
    a("[CIDADE]/[UF] | Município e Estado da sede ou do ato")
    a("[NOME COMPLETO] e [QUALIFICAÇÃO] | Nome civil completo, sem abreviações; qualificação completa quando pedida")
    a("[NÚMERO] | Número do documento, do registro, do CPF, do RG ou da matrícula, conforme o contexto")
    a("[VALOR] e [VALOR POR EXTENSO] | Valor em reais, com número e extenso")
    a("[ENDEREÇO COMPLETO, COM CEP] | Logradouro, número, complemento, bairro, cidade, UF e CEP")
    a("[NOME DO PRESIDENTE] e [NOME DO PASTOR TITULAR] | Ocupantes atuais dos cargos, conforme a última ata de eleição")
    a("%ENDGRID")
    a("## 5. ORDEM DE UTILIZAÇÃO")
    a("I – Fundação: Edital de Convocação → Assembleia de Fundação (aprovação do Estatuto e da Declaração de Fé, eleições e posse) → Ata de Fundação.")
    a("II – Registro: visto de advogado → Requerimento de Registro no RCPJ → CNPJ com natureza jurídica 322-0 (Organização Religiosa) → pedidos de imunidade de IPTU, ITBI, IPVA e ITCMD → contas bancárias, contador e alvarás.")
    a("III – Primeiros 180 dias: aprovação do Regimento Interno pela Assembleia; aprovação do Código de Conduta e das Políticas pela Diretoria; designação do encarregado de dados e do(a) Coordenador(a) de Proteção; contrato do imóvel do templo (comodato ou locação).")
    a("IV – Funcionamento contínuo: atas mensais da Diretoria e do Conselho Ministerial, atas trimestrais do Conselho Fiscal, recibos, fichas de membros, certificados, contratos de pessoas e serviços.")
    a("V – Todo ano: parecer do Conselho Fiscal, relatório anual de atividades e Assembleia Geral Ordinária até 30 de abril.")
    a("## 6. FICHAS DOS DOCUMENTOS")
    n = 0
    for pasta, desc, chaves in PASTAS:
        a(f"## Pasta {pasta.replace('_', ' ')}")
        for k in chaves:
            n += 1
            d = DOCUMENTOS[k]
            src, out = chave_arquivo[k]
            itens, n_art = estrutura(src)
            cps = campos(src)
            especificos = [c for c, _ in cps.most_common() if c not in COMUNS]
            a(f"### 6.{n}. {d['titulo']}")
            a("%GRID 5,11")
            a("Item | Descrição")
            a(f"Arquivo | {pasta}/{n_idx(pasta, k)}_{NOME_ENVIO[k]}.docx")
            a(f"Finalidade | {d['finalidade']}")
            a(f"Base legal | {d['base']}")
            a(f"Quem prepara, aprova e assina | {d['quem']}")
            a(f"Guarda e prazos | {d['guarda']}")
            a("%ENDGRID")
            if itens or n_art:
                partes = []
                if n_art:
                    partes.append(f"{n_art} artigos")
                a(f"P Estrutura{' (' + ', '.join(partes) + ')' if partes else ''}: " + "; ".join(itens) + "." if itens else f"P Estrutura: {n_art} artigos numerados.")
            a("P Regras de preenchimento:")
            for i, r in enumerate(d["regras"], 1):
                a(f"{roman(i)} – {r}")
            if especificos:
                a("P Campos específicos a preencher (além dos campos comuns da seção 4): " + "; ".join(f"[{c}]" for c in especificos[:60]) + ("; e outros de menor recorrência." if len(especificos) > 60 else "."))
            a(f"P Cuidados: {d['cuidados']}")
    a("## 7. PASTAS DE ENVIO AO PASTOR")
    a("P Os arquivos foram organizados em pastas numeradas, na ordem em que serão utilizados:")
    a("%GRID 6,10")
    a("Pasta | Conteúdo")
    a("00_Leia_Primeiro | Este manual e a Nota Técnica")
    for pasta, desc, chaves in PASTAS:
        a(f"{pasta} | {desc} " + ", ".join(DOCUMENTOS[k]['titulo'] for k in chaves) + ".")
    a("%ENDGRID")
    a("P Não fazem parte do envio os arquivos de texto (.txt) e o programa gerador, que servem apenas para manter e atualizar os modelos.")
    (base / "manual.txt").write_text("\n".join(L) + "\n", encoding="utf-8")


def roman(i):
    return ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"][i - 1]


def n_idx(pasta, k):
    for p, _, chaves in PASTAS:
        if p == pasta:
            return f"{chaves.index(k) + 1:02d}"


def nota_docx():
    """Converte a NOTA_TECNICA.md em .docx (títulos, listas, tabelas e negrito)."""
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21)
    sec.top_margin = sec.left_margin = Cm(3)
    sec.bottom_margin = sec.right_margin = Cm(2)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Times New Roman", Pt(11)

    def runs(p, text):
        for i, part in enumerate(re.split(r"\*\*", text)):
            r = p.add_run(part.replace("`", ""))
            r.bold = i % 2 == 1

    lines = (base / "NOTA_TECNICA.md").read_text(encoding="utf-8").splitlines()
    i = 0
    while i < len(lines):
        l = lines[i]
        if l.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                if not re.match(r"^\|[-| ]+\|$", lines[i]):
                    rows.append([c.strip() for c in lines[i].strip("|").split("|")])
                i += 1
            t = doc.add_table(rows=len(rows), cols=len(rows[0]))
            t.style = "Table Grid"
            for r, row in enumerate(rows):
                for c, val in enumerate(row):
                    cell = t.rows[r].cells[c]
                    cell.text = ""
                    runs(cell.paragraphs[0], val)
            continue
        if l.startswith("# "):
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(l[2:]); r.bold = True; r.font.size = Pt(14)
        elif l.startswith("## "):
            p = doc.add_paragraph(); r = p.add_run(l[3:]); r.bold = True; r.font.size = Pt(12)
            p.paragraph_format.space_before = Pt(10)
        elif re.match(r"^\s*[-*] ", l):
            indent = (len(l) - len(l.lstrip())) // 2
            p = doc.add_paragraph(style="List Bullet" if indent == 0 else "List Bullet 2")
            runs(p, re.sub(r"^\s*[-*] ", "", l))
        elif re.match(r"^\s*\d+\. ", l):
            p = doc.add_paragraph(style="List Number" if not l.startswith(" ") else "List Number 2")
            runs(p, re.sub(r"^\s*\d+\. ", "", l))
        elif l.strip():
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            runs(p, l)
        i += 1
    out = base / "Nota_Tecnica_Igreja_IBP.docx"
    doc.save(out)
    return out


def montar_envio():
    if ENVIO.exists():
        shutil.rmtree(ENVIO)
    leia = ENVIO / "00_Leia_Primeiro"
    leia.mkdir(parents=True)
    shutil.copy(base / "Manual_dos_Documentos_Igreja_IBP.docx", leia / "01_Manual_dos_Documentos.docx")
    shutil.copy(base / "Nota_Tecnica_Igreja_IBP.docx", leia / "02_Nota_Tecnica.docx")
    out_de = {Path(s).stem: o for s, o in DOCS}
    for pasta, _, chaves in PASTAS:
        d = ENVIO / pasta
        d.mkdir()
        for i, k in enumerate(chaves, 1):
            shutil.copy(base / out_de[k], d / f"{i:02d}_{NOME_ENVIO[k]}.docx")
    return sum(1 for _ in ENVIO.rglob("*.docx"))


def script_windows():
    """Gera o script PowerShell que organiza a pasta local do usuário na mesma estrutura."""
    out_de = {Path(s).stem: o for s, o in DOCS}
    linhas = []
    for pasta, _, chaves in PASTAS:
        for i, k in enumerate(chaves, 1):
            padrao = re.sub(r"_Igreja_(ISAC|IBP)(_v2)?$", "", Path(out_de[k]).stem)  # aceita nomes antigos e novos
            linhas.append(f'    @{{ Padrao = "{padrao}"; Pasta = "{pasta}"; Nome = "{i:02d}_{NOME_ENVIO[k]}.docx" }}')
    linhas.insert(0, '    @{ Padrao = "Nota_Tecnica"; Pasta = "00_Leia_Primeiro"; Nome = "02_Nota_Tecnica.docx" }')
    linhas.insert(0, '    @{ Padrao = "Manual_dos_Documentos"; Pasta = "00_Leia_Primeiro"; Nome = "01_Manual_dos_Documentos.docx" }')
    ps = f"""# Organiza os documentos da igreja em pastas numeradas para envio ao Pastor Isac.
# Uso (PowerShell): clique com o botão direito no arquivo > "Executar com o PowerShell",
# ou abra o PowerShell e rode:  powershell -ExecutionPolicy Bypass -File .\organizar_envio_pastor.ps1
# O script SOMENTE COPIA arquivos: nada é apagado, movido ou alterado na pasta original.
param(
    [string]$Origem = (Join-Path $env:USERPROFILE "Downloads\Documentos - Constituição Igreja IBP")
)

$Destino = Join-Path $Origem "Envio - Pastor Isac"
$Mapa = @(
{chr(10).join(linhas)}
)

if (-not (Test-Path -LiteralPath $Origem)) {{
    Write-Host "Pasta de origem não encontrada: $Origem" -ForegroundColor Red
    Write-Host "Rode novamente informando a pasta: .\organizar_envio_pastor.ps1 -Origem 'C:\caminho\da\pasta'"
    exit 1
}}

$arquivos = Get-ChildItem -LiteralPath $Origem -File -Filter *.docx
$usados = @{{}}
$faltando = @()
foreach ($item in $Mapa) {{
    # aceita variações como "Nome (1).docx"; se houver mais de uma, usa a mais recente
    $candidatos = $arquivos | Where-Object {{ $_.BaseName -like "*$($item.Padrao)*" }} | Sort-Object LastWriteTime -Descending
    if ($candidatos) {{
        $arq = $candidatos[0]
        $pasta = Join-Path $Destino $item.Pasta
        New-Item -ItemType Directory -Force -Path $pasta | Out-Null
        Copy-Item -LiteralPath $arq.FullName -Destination (Join-Path $pasta $item.Nome) -Force
        $candidatos | ForEach-Object {{ $usados[$_.FullName] = $true }}
        Write-Host "OK   $($item.Pasta)\$($item.Nome)   <=  $($arq.Name)"
    }} else {{
        $faltando += "$($item.Pasta)\$($item.Nome)  (procure um arquivo com '$($item.Padrao)' no nome)"
    }}
}}

Write-Host ""
if ($faltando.Count -gt 0) {{
    Write-Host "Documentos NÃO encontrados na pasta de origem (baixe-os e rode de novo):" -ForegroundColor Yellow
    $faltando | ForEach-Object {{ Write-Host "  - $_" }}
}}
$sobras = $arquivos | Where-Object {{ -not $usados.ContainsKey($_.FullName) }}
if ($sobras) {{
    Write-Host ""
    Write-Host "Arquivos da pasta que NÃO fazem parte do envio (confira se devem ir ao Pastor):" -ForegroundColor Cyan
    $sobras | ForEach-Object {{ Write-Host "  - $($_.Name)" }}
}}
Write-Host ""
Write-Host "Concluído. Pasta de envio: $Destino" -ForegroundColor Green
"""
    out = base.parent / "envio-pastor-isac" / "organizar_envio_pastor.ps1"
    out.write_text(ps, encoding="utf-8-sig")
    return out


if __name__ == "__main__":
    assert set(DOCUMENTOS) == {Path(s).stem for s, _ in DOCS}, "DOCUMENTOS e DOCS divergem"
    assert set(NOME_ENVIO) == set(DOCUMENTOS)
    assert sorted(k for _, _, c in PASTAS for k in c) == sorted(DOCUMENTOS), "PASTAS incompleto"
    escrever_manual()
    build("manual.txt", "Manual_dos_Documentos_Igreja_IBP.docx")
    print(nota_docx().name)
    print(f"{montar_envio()} arquivos em {ENVIO.name}/")
    print(script_windows().name)
