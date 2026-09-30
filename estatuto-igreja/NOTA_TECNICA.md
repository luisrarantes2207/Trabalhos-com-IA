# Nota técnica: novo Estatuto Social da Igreja (organização religiosa)

Arquivos:
- `Estatuto_Social_Igreja_ISAC_v2.docx`: o estatuto pronto para preencher e registrar.
- `Ata_Assembleia_Fundacao_Igreja_ISAC.docx`: ata da Assembleia de Fundação, com o Anexo A (lista de presença dos fundadores) e o Anexo B (termo de posse e declaração de desimpedimento, um por eleito).
- `Edital_Convocacao_Fundacao_Igreja_ISAC.docx`: edital de convocação da Assembleia de Fundação, com a certidão de afixação e divulgação.
- `Requerimento_Registro_RCPJ_Igreja_ISAC.docx`: requerimento de registro ao Registro Civil de Pessoas Jurídicas, com a lista de documentos que o acompanham.
- `estatuto.txt`, `ata_fundacao.txt`, `edital_convocacao.txt` e `requerimento_registro.txt`: o texto-fonte. Edite este arquivo e rode `python3 gerar_docx.py` para gerar os dois .docx de novo. Os artigos do estatuto são renumerados automaticamente, e as remissões da ata acompanham essa numeração.

## 1. Base legal verificada (setembro de 2026)

| Tema | Dispositivo |
|---|---|
| Natureza jurídica e autonomia | CC, art. 44, IV, §§ 1º e 2º; CF, arts. 5º, VI, e 19, I |
| Conteúdo obrigatório do estatuto | CC, arts. 46 e 54; Lei 6.015/1973, arts. 114 e 120 |
| Assembleia: competência privativa, convocação por 1/5 | CC, arts. 59 e 60 |
| Exclusão de membro com defesa e recurso | CC, art. 57 |
| Assembleia eletrônica | CC, art. 48-A |
| Destino do patrimônio na dissolução | CC, art. 61 |
| Imunidade de impostos | CF, art. 150, VI, "b", e § 4º, na redação da EC 132/2023 ("entidades religiosas e templos de qualquer culto, inclusive suas organizações assistenciais e beneficentes") |
| IPTU de imóvel alugado usado como templo | CF, art. 156, § 1º-A (EC 116/2022) |
| IBS/CBS (reforma tributária) | CF, art. 149-B, parágrafo único; LC 214/2025, art. 9º, II, e § 2º (define a entidade religiosa como a PJ sem fins lucrativos que tem como objetivos "professar a fé religiosa e praticar a religião"). O § 4º diz que a imunidade **não** alcança as aquisições da igreja. |
| ITCMD sobre doações e heranças recebidas | LC 227/2026, art. 149, I, "c", e § 1º, II: a imunidade só vale para bens relacionados às finalidades essenciais |
| CTN | art. 9º, IV, "b" (redação da LC 214/2025) e art. 14 (requisitos adotados voluntariamente no art. 5º, § 1º) |
| Prebenda/côngrua de ministros | Lei 8.212/1991, art. 22, §§ 13, 14 e 16 |
| Voluntariado | Lei 9.608/1998 |
| Visto de advogado no registro | Lei 8.906/1994, art. 1º, § 2º |

## 2. O que mudou em relação ao rascunho original e por quê

1. **A finalidade essencial está redigida com as palavras da LC 214/2025** ("professar a fé religiosa e praticar a religião"), no art. 4º. Isso enquadra a igreja diretamente na definição legal de entidade imune ao IBS e à CBS.
2. **Atividades-meio** (cantina, livraria, eventos, aluguel de bem ocioso): ficam permitidas, desde que o resultado seja aplicado integralmente nas finalidades essenciais e a contabilidade seja segregada (art. 4º, §§ 1º e 2º). Esse é o critério do art. 150, § 4º, da CF, e é o ponto em que a fiscalização mais contesta a imunidade.
3. **Requisitos do art. 14 do CTN adotados voluntariamente** (art. 5º, § 1º). Hoje a lei não os exige de templos, mas o Fisco e os municípios costumam cobrá-los na prática, e municípios frequentemente pedem o estatuto para reconhecer a imunidade.
4. **Cláusulas que não podem ser alteradas** (art. 20, § 3º): a natureza religiosa e não lucrativa, a proibição de distribuir patrimônio ou renda e o destino do patrimônio não podem ser mudados por reforma estatutária.
5. **Quóruns qualificados com quórum mínimo de instalação**: o original exigia 2/3 "dos presentes" sem mínimo de presença, o que permite decisões com muito poucos membros. Agora são 2/3 para reforma do estatuto, destituições e venda de imóveis, e 3/4 para mudar a Declaração de Fé, dissolver a igreja ou filiá-la/desfiliá-la de convenção.
6. **Conselho Fiscal obrigatório**, com regra contra nepotismo. No original ele era opcional.
7. **Controles financeiros**: duas assinaturas para movimentar dinheiro, contagem de ofertas por duas pessoas, PIX e cartões só no CNPJ da igreja, bens só em nome da igreja, política de conflito de interesses, proibição de empréstimos, avais e fianças.
8. **Sustento pastoral separado da remuneração de diretoria** (art. 34). Isso protege a imunidade e o enquadramento previdenciário. A prebenda é proibida em percentual da arrecadação, porque esse formato caracteriza distribuição disfarçada de renda e remuneração vinculada à produção, o que é vedado pelo § 13 do art. 22 da Lei 8.212.
9. **Devido processo para exclusão de membros**, com defesa e recurso à Assembleia (CC, art. 57), preservando a autonomia doutrinária da igreja (art. 3º).
10. **Congregações/filiais** sem personalidade jurídica própria. Os bens das congregações pertencem à igreja, e uma congregação que se separar não leva patrimônio.
11. **LGPD** (a convicção religiosa é dado sensível), **proteção de crianças e vulneráveis** e **proibição de uso político-eleitoral**.
12. **Anexo I – Declaração de Fé**, redigida (14 seções), protegida por quórum de 3/4. Segue a linha cristã evangélica e tem alternativas entre colchetes para o parágrafo pentecostal/carismático, a forma do batismo e quem pode participar da Ceia. A seção IX (casamento) inclui uma cláusula que permite à igreja recusar celebrações contrárias à sua fé, com respeito às pessoas e à lei civil. A seção XIV liga a Declaração ao Estatuto: ela é interpretada pelo Conselho Ministerial, com recurso à Assembleia.

## 3. Campos a preencher (entre [COLCHETES])

Nome e sigla, endereço, prazo de carência para membro votante, prazos de mandato e número de reconduções, a opção sobre o Pastor Titular (membro nato ou eleito como Presidente; mandato fixo ou prazo indeterminado), o limite de endividamento (% da receita), o valor de receita que obriga à auditoria, o número mínimo de membros que impede a dissolução, a entidade que recebe o patrimônio na dissolução, e as opções doutrinárias da Declaração de Fé (seção IV, parágrafo pentecostal; seção VIII, batismo e Ceia).

## 4. Checklist após a aprovação

1. Ata da Assembleia (de fundação ou de reforma), com lista de presença e qualificação completa dos eleitos.
2. Visto de advogado e registro no Registro Civil de Pessoas Jurídicas.
3. CNPJ com natureza jurídica **322-0 – Organização Religiosa** e inscrição de cada congregação como filial.
4. Pedido de reconhecimento de imunidade de **IPTU/ITBI** na prefeitura (inclusive para imóvel alugado, com base na CF, art. 156, § 1º-A) e de **IPVA/ITCMD** no Estado, conforme a legislação local.
5. Contador habilitado, escrituração conforme a ITG 2002 (R1) e cumprimento das obrigações acessórias (ECF, EFD-Reinf/DCTFWeb, eSocial se houver empregados, e ECD quando obrigatória).
6. Regimento Interno, código de conduta, política de proteção de menores e política de privacidade.

## 5. Pontos de atenção (riscos que o estatuto sozinho não resolve)

- **A imunidade depende da prática, não só do texto.** Usar bens ou recursos para fins pessoais, pagar prebendas desproporcionais ou manter atividade comercial preponderante pode levar à perda da imunidade e à responsabilização pessoal dos dirigentes (CC, art. 50; CTN, art. 135).
- **A imunidade não cobre o que a igreja compra.** O ICMS e o IBS/CBS embutidos no preço das aquisições continuam sendo pagos (LC 214/2025, art. 9º, § 4º).
- **Taxas e contribuições não são imunes.** Isso inclui a contribuição previdenciária sobre empregados, que continua devida.
- **Revisão profissional.** Este texto é uma minuta técnica e deve ser revisado por advogado inscrito na OAB antes do registro, com atenção especial às exigências do cartório e da prefeitura do município-sede.

## 6. Ata de fundação: cuidados para o registro

- **Regra de transição (art. 50, § 1º, do Estatuto):** sem ela, nenhum fundador cumpriria a carência de 6 meses para votar (art. 9º, I) nem a de 2 anos para ser eleito (art. 24, I). A regra dispensa esses prazos somente na fundação. Por segurança, os fundadores devem ser pessoas civilmente capazes. Menores podem ser admitidos como membros depois.
- **Qualificação completa de cada eleito** (nome, nacionalidade, estado civil, profissão, RG, CPF e endereço): é exigida pelo art. 120 da Lei 6.015/1973, e sua falta é o motivo mais comum de o cartório devolver o pedido.
- **Assinaturas e visto:** a ata é assinada pelo presidente e pelo secretário da Assembleia e pelos eleitos, com visto de advogado. Todas as folhas do Estatuto devem ser rubricadas. Consulte o cartório local sobre reconhecimento de firma e número de vias.
- **Preenchimento dos campos:** preencha todos os [COLCHETES] e apague as alternativas que não se aplicam. Não deixe espaços em branco nem rasuras. O resultado de cada votação deve bater com o número de presentes.
- **Documentos para o cartório (normalmente):** requerimento assinado pelo Presidente (`Requerimento_Registro_RCPJ_Igreja_ISAC.docx`), uma via do Estatuto e da ata em papel ou em meio eletrônico (Lei 6.015, art. 121, na redação atual; alguns cartórios ainda pedem mais vias), lista de presença, cópia do edital de convocação e documentos de identidade dos dirigentes. Confirme a lista exata com o Registro Civil de Pessoas Jurídicas da comarca.

## 7. Edital e requerimento

- **Edital:** antes do registro a igreja ainda não existe, então quem convoca é uma Comissão Organizadora. O prazo de 15 dias repete a regra do Estatuto para assembleias extraordinárias (art. 18, § 2º). A ordem do dia é idêntica à da ata. Guarde as provas da divulgação (a certidão de afixação e capturas de tela), porque a ata faz referência ao edital.
- **Requerimento:** fundamenta o pedido no CC, arts. 44 e 45, e na Lei 6.015/1973, arts. 114, 120 e 121. Ele aponta ao cartório onde está cada informação exigida pelo art. 120 e identifica o apresentante (art. 120, VI). Pela redação atual do art. 121, § 1º, o requerimento pode ser dispensado quando o representante legal assinou o estatuto. Mesmo assim, vale apresentá-lo, porque ele organiza o pedido e informa os contatos para eventuais exigências.
- **Antes de protocolar:** confira no próprio cartório as normas de serviço da Corregedoria do seu Estado (número de vias, reconhecimento de firma e documentos adicionais) e faça uma pesquisa de nome para evitar homonímia.
