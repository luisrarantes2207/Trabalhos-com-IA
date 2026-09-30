# Nota técnica: novo Estatuto Social da Igreja (organização religiosa)

Arquivos:
- `Estatuto_Social_Igreja_ISAC_v2.docx`: o estatuto pronto para preencher e registrar.
- `Ata_Assembleia_Fundacao_Igreja_ISAC.docx`: ata da Assembleia de Fundação, com o Anexo A (lista de presença dos fundadores) e o Anexo B (termo de posse e declaração de desimpedimento, um por eleito).
- `Edital_Convocacao_Fundacao_Igreja_ISAC.docx`: edital de convocação da Assembleia de Fundação, com a certidão de afixação e divulgação.
- `Requerimento_Registro_RCPJ_Igreja_ISAC.docx`: requerimento de registro ao Registro Civil de Pessoas Jurídicas, com a lista de documentos que o acompanham.
- `Regimento_Interno_Igreja_ISAC.docx`: Regimento Interno (53 artigos), a ser aprovado pela Assembleia Geral em até 180 dias da fundação.
- `Codigo_de_Conduta_Igreja_ISAC.docx` e `Politica_Protecao_Menores_Vulneraveis_Igreja_ISAC.docx`: normas complementares aprovadas pela Diretoria Executiva, ouvido o Conselho Ministerial (Estatuto, art. 43).
- `Politica_Privacidade_Protecao_Dados_Igreja_ISAC.docx`: política de privacidade (aviso público aos titulares e normas internas), aprovada pela Diretoria Executiva (Estatuto, art. 43, IV).
- `Termo_Adesao_Servico_Voluntario_Igreja_ISAC.docx`: termo de adesão ao serviço voluntário (Lei 9.608/1998), com registro de atividades e termo de encerramento.
- `Contrato_Comodato_Templo_Igreja_ISAC.docx`: modelo de contrato de comodato do imóvel usado como templo, com laudo de vistoria e termo de devolução.
- `Contrato_Locacao_Templo_Igreja_ISAC.docx`: modelo de contrato de locação não residencial do imóvel usado como templo, com laudo de vistoria e termo de devolução.
- `Recibo_Dizimos_Ofertas_Igreja_ISAC.docx`: instruções à Tesouraria e modelos de recibo de contribuição (com canhoto), de declaração anual de contribuições e de requerimento de devolução por erro material.
- `Termo_Doacao_Bens_Igreja_ISAC.docx`: termo de doação de bens móveis, inclusive veículos, com roteiro para a doação de imóveis.
- Arquivos `.txt` com o texto-fonte de cada documento: o texto-fonte. Edite este arquivo e rode `python3 gerar_docx.py` para gerar os dois .docx de novo. Os artigos do estatuto são renumerados automaticamente, e as remissões dos demais documentos acompanham essa numeração. Cada documento normativo tem numeração própria: `[[chave]]` remete ao Estatuto, `<<chave>>` ao Regimento Interno, `{{chave}}` ao próprio documento `{{arquivo:chave}}` a outro documento (por exemplo, `{{codigo_conduta:cons}}`) e `((chave))` a cláusula do próprio termo.

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

## 8. Regimento Interno

Ele regulamenta todos os pontos que o Estatuto delega ao Regimento: modos de admissão (art. 8º), faltas que caracterizam conduta incompatível (art. 13), eleição da Diretoria por chapa (art. 23), composição do Conselho Ministerial (art. 32), limite do fundo fixo de caixa (art. 38), alçada de despesas autorizadas por um só dirigente (art. 41), integridade (art. 43) e mediação interna (art. 48).

**Destaques:**
- **Disciplina:** as medidas são graduais, da admoestação reservada à exclusão, e a exposição pública do membro é proibida (reduz o risco de ação por dano moral). O procedimento prevê impedimentos de quem julga, direito a acompanhante ou advogado, afastamento cautelar em casos graves, e recurso à Assembleia em sessão reservada com voto secreto, sem possibilidade de agravar a punição.
- **Proteção de menores:** regra de pelo menos dois adultos em cada atividade, certidões criminais para voluntários e comunicação obrigatória ao Conselho Tutelar (ECA, arts. 13 e 70-B; o parágrafo único do art. 70-B alcança expressamente quem exerce "ministério"). Nenhum órgão da igreja pode impedir ou atrasar essa comunicação.
- **Finanças:** alçadas de despesa, exigência de 3 orçamentos, proibição de fracionar despesas, contagem de ofertas em duplicidade com depósito em até 2 dias úteis, proibição de PIX em contas pessoais, e ministérios e congregações sem caixa próprio.
- **Sustento pastoral:** proibidas metas, comissões ou qualquer vínculo com a arrecadação (Lei 8.212, art. 22, § 13).
- **Eleições e conduta:** Comissão Eleitoral independente e termo de transição entre diretorias. Proibição de uso político-eleitoral da igreja, pois a legislação eleitoral trata templos como bens de uso comum. O dispositivo foi confirmado: Lei 9.504/1997, art. 37, § 4º.
- **Valores a preencher:** alçadas em R$, fundo fixo de caixa, prazos e número de presbíteros ou oficiais, conforme o porte da igreja.

## 9. Código de Conduta e Política de Proteção

Os dois são aprovados pela Diretoria Executiva, ouvido o Conselho Ministerial, sem precisar de Assembleia (Estatuto, art. 43). Todo colaborador assina o termo anexo a cada um.

**Código de Conduta (22 artigos):**
- **Abrangência e condutas vedadas:** vale para dirigentes, ministros, líderes, voluntários, empregados e prestadores de serviço. Proíbe assédio moral e sexual, discriminação e retaliação.
- **Autoridade espiritual:** proíbe coagir pessoas a contribuir e prometer bênçãos em troca de dinheiro, coerente com a Declaração de Fé, seção VI. Isso protege o caráter voluntário das contribuições, do qual depende a imunidade.
- **Aconselhamento:** em ambiente visível, com encaminhamento a profissionais quando necessário e proibição de relacionamento afetivo com a pessoa aconselhada.
- **Recursos:** conflito de interesses, presentes de fornecedores e anticorrupção (Lei 12.846/2013).
- **Comunicação e política:** regras para redes sociais e neutralidade eleitoral. O colaborador que se candidatar se afasta da liderança durante a campanha; essa regra é opcional e pode ser retirada.
- **Consequências por vínculo:** disciplina eclesiástica para membros, legislação trabalhista para empregados, desligamento para voluntários e rescisão para prestadores.

**Política de Proteção (24 artigos e 4 anexos):**
- **Governança:** um(a) Coordenador(a) de Proteção, preferencialmente independente do Pastor Titular.
- **Habilitação de colaboradores:** certidões criminais renovadas a cada 12 meses, ou a cada 6 meses se a igreja receber recursos públicos para atividades com crianças (ECA, art. 59-A, conferido); impedimentos; capacitação; e período supervisionado.
- **Regras operacionais:** regra dos dois adultos, proporção de crianças por adulto, entrada e saída com senha, banheiros, contato físico, comunicação digital, transporte, viagens (ECA, art. 83, conferido), medicamentos e imagem (LGPD, art. 14).
- **Resposta a revelações:** acolher sem investigar e comunicar imediatamente ao Conselho Tutelar (ECA, arts. 13 e 70-B; Lei 13.431/2017, art. 13, conferido). Nenhum órgão da igreja pode impedir essa comunicação, é proibida qualquer mediação entre a vítima e o suspeito, e o suspeito é afastado imediatamente.
- **Pessoas com histórico de violência:** podem assistir aos cultos mediante acordo de participação supervisionada.
- **Anexos:** termo do colaborador, ficha e autorizações da criança (com consentimento conforme a LGPD), autorização para viagem e formulário sigiloso de registro de incidente.
- **Adaptações locais:** a proporção de crianças por adulto e os prazos entre colchetes são sugestões; ajuste ao espaço e ao número de voluntários. Nas viagens, confira a resolução vigente do CNJ sobre autorização de viagem.

## 10. Política de Privacidade e Proteção de Dados

**Estrutura:**
- **Parte I (aviso aos titulares):** controlador e encarregado; tabela com dados, finalidades e base legal para cada público (membros, visitantes, crianças, doadores, colaboradores, empregados, dirigentes, assistidos, participantes de eventos, pessoas filmadas e usuários do site); compartilhamentos; transferência internacional; tabela de prazos de conservação; direitos do titular (LGPD, art. 18, com resposta em até 15 dias conforme o art. 19); segurança; e cookies.
- **Parte II (normas internas):** papéis de cada um; registro das operações e relatório de impacto (LGPD, arts. 37 e 38); regras do dia a dia (acessos, credenciais, planilhas, papéis, grupos de mensagens, pedidos de oração); contratos com fornecedores que tratam dados; plano de resposta a incidentes; e relatório anual.
- **Anexos:** termo de consentimento do membro, formulário de solicitação do titular e cláusula de proteção de dados para contratos com fornecedores.

**Escolhas de fundo:**
- **Religião como dado sensível:** estar nos registros da igreja já revela a religião da pessoa (LGPD, art. 5º, II). Por isso, os dados de membros se baseiam em consentimento específico (art. 11, I). Os registros exigidos por lei ou pelo Estatuto se baseiam em obrigação legal ou exercício regular de direitos (art. 11, II, "a" e "d"), e continuam guardados mesmo se o consentimento for revogado. O legítimo interesse nunca é usado para dados sensíveis.
- **Aconselhamento pastoral:** a orientação é não registrar detalhes por escrito e não incluí-los em cadastros gerais.
- **Transmissões de cultos:** aviso na entrada, área fora das câmeras e enquadramento no púlpito, sem focar pessoas do público.
- **Encarregado obrigatório:** as organizações sem fins lucrativos podem ser agentes de pequeno porte (Resolução CD/ANPD nº 2/2022), mas esse regime não vale para tratamento de alto risco. Dados de religião e de crianças podem se enquadrar nesse critério, então a Política mantém o encarregado de qualquer forma.
- **Incidentes:** comunicação à ANPD e aos titulares em 3 dias úteis, ou no prazo diferenciado de pequeno porte, se aplicável (Resolução CD/ANPD nº 15/2024, conferida por pesquisa na web).
- **Adaptações:** os prazos de conservação entre colchetes e o prazo de 48 horas para fornecedores comunicarem incidentes são sugestões. Ajuste a tabela de dados às atividades que a igreja realmente tem, por exemplo retirando site ou aplicativo se não existirem.

## 11. Termo de adesão ao serviço voluntário

**Base legal:** Lei 9.608/1998. O termo de adesão escrito, com o objeto e as condições do serviço, é o que a lei exige (art. 2º). O serviço não gera vínculo empregatício nem obrigação trabalhista ou previdenciária (art. 1º, parágrafo único). Só se ressarcem despesas comprovadas e previamente autorizadas (art. 3º). O Estatuto remete a essa lei no art. 35. Esses artigos foram citados sem conferência do texto, porque a base de legislação estava sem consultas, então vale o advogado checar.

**Como evitar o reconhecimento de vínculo de emprego.** O termo sozinho não impede uma condenação trabalhista: a Justiça do Trabalho olha a prática. Por isso o termo:
- trata a disponibilidade como uma referência definida pelo próprio voluntário, sem controle de jornada nem penalidade por falta;
- proíbe ajuda de custo fixa ou periódica sem comprovante, que costuma ser vista como salário disfarçado;
- deixa claro que ser voluntário não é condição para ser membro.

Na prática, evite também metas, escalas impostas, advertências em estilo trabalhista e pagamentos habituais. Se alguém trabalha com habitualidade, horário fixo e subordinação, contrate pela CLT.

**Outras cláusulas:**
- **Crianças e vulneráveis:** o voluntário que atuar com esses públicos só começa depois de cumprir todos os requisitos da Política de Proteção.
- **Imagem e obras:** autorização de imagem revogável e licença gratuita, não exclusiva, das obras produzidas para a igreja (músicas, artes, vídeos), preservados os direitos morais do autor (Lei 9.610/1998).
- **Dados pessoais:** tratamento conforme a Política de Privacidade.
- **Encerramento:** qualquer das partes pode encerrar a qualquer tempo, com devolução de bens e credenciais.
- **Adolescentes de 14 a 17 anos:** cláusula opcional que exige a assinatura dos pais, compatibilidade com a escola e proíbe trabalho noturno ou perigoso.

**Anexos:** registro de atividades e ressarcimentos, que serve de base para a declaração de serviço voluntário, e termo de encerramento com quitação.

## 12. Contrato de comodato do templo

**Base legal:** CC, arts. 579 a 585. O comodato é o empréstimo gratuito de um bem. O comodante não pode retomar o imóvel antes do prazo, salvo necessidade urgente reconhecida pelo juiz (art. 581). O comodatário deve conservar o bem como se fosse seu e, se atrasar a devolução, paga aluguel arbitrado (art. 582). Ele também não recupera as despesas de uso (art. 584). Esses artigos foram citados sem conferência do texto, porque a base de legislação estava sem consultas.

**Pontos de atenção:**
- **IPTU:** a Constituição (art. 156, § 1º-A) afasta o IPTU de templos em imóvel de terceiro quando a igreja é **locatária**, e não fala em comodato. O reconhecimento da imunidade para imóvel emprestado é incerto e depende do município. O contrato prevê um pedido conjunto de imunidade ou isenção e define quem paga se o pedido for negado. Não transforme o comodato em "locação com aluguel simbólico" só para buscar a imunidade: isso pode ser visto como simulação. Avalie com o advogado e verifique se a lei municipal prevê isenção.
- **Prazo determinado:** recomendado (a sugestão é 10 anos). O prazo protege a igreja contra uma retomada repentina, e o contrato ainda exige aviso prévio de 180 dias para a devolução ao fim do prazo, para dar tempo de transferir o templo.
- **Venda do imóvel:** o comodato não é direito real e, em regra, não obriga quem compra o imóvel. Por isso o comodante se obriga a exigir do comprador que respeite o contrato, sob pena de indenizar a igreja. O registro no Registro de Títulos e Documentos dá prova da data do contrato. Para uma proteção real contra terceiros, a alternativa seria a locação com cláusula de vigência averbada na matrícula (Lei 8.245/1991, art. 8º).
- **Comodante ligado à igreja:** se o dono do imóvel for dirigente, ministro ou parente de um deles, aplica-se o Estatuto, art. 40. É preciso parecer do Conselho Fiscal e aprovação da Diretoria sem a participação do interessado, e o contrato declara que o comodante não ganha nenhum direito sobre a igreja.
- **Benfeitorias:** adaptações exigidas pelo Poder Público (Bombeiros e acessibilidade) podem ser feitas sem autorização. Bancos, som e outros bens removíveis continuam da igreja. A indenização de obras fica como opção a escolher.
- **Assinatura do cônjuge:** a lei não exige a assinatura do cônjuge do comodante no comodato (CC, art. 1.647), mas o contrato inclui o cônjuge como anuente para prevenir disputas.

## 13. Contrato de locação do templo

**Base legal:** Lei 8.245/1991 (Lei do Inquilinato). Os artigos citados abaixo não foram conferidos no texto oficial, porque a base de legislação estava sem consultas. Confira antes de usar.

**Proteções específicas para igrejas:**
- **Art. 53:** para imóvel usado por entidade religiosa devidamente registrada, o locador só pode rescindir a locação nas hipóteses do art. 9º (acordo, infração, falta de pagamento, reparos urgentes exigidos pelo Poder Público) ou para demolição, edificação licenciada ou reforma que aumente a área útil em pelo menos 50%. Ele não pode pedir o imóvel sem motivo, nem depois de vencido o prazo. Por isso o contrato exige que a igreja esteja registrada e informe o número do registro.
- **Art. 63, § 3º:** em caso de despejo, a entidade religiosa tem prazo de um ano para desocupar, ou de seis meses se entre a citação e a sentença já tiver passado mais de um ano.
- **IPTU:** a CF, art. 156, § 1º-A (EC 116/2022), afasta expressamente o IPTU de templos em imóvel alugado. O contrato prevê o pedido de reconhecimento à prefeitura e quem paga enquanto ele não sai.

**Outros cuidados:**
- **Venda do imóvel:** a cláusula de vigência e o direito de preferência (arts. 8º e 27 a 34) só valem contra o comprador se o contrato estiver registrado ou averbado na matrícula. A lei exige averbação pelo menos 30 dias antes da venda para a preferência (art. 33). A igreja deve providenciar isso logo após a assinatura.
- **Tributos:** quando o locador é pessoa física, a igreja, como pessoa jurídica, retém o imposto de renda na fonte sobre o aluguel. Ser imune não a dispensa de reter tributos devidos por terceiros. Depois da reforma tributária, se o locador for contribuinte do IBS e da CBS, esses tributos estão incluídos no aluguel, e a imunidade da igreja não alcança o que ela paga a terceiros (LC 214/2025, art. 9º, § 4º). Confirme com o contador.
- **Garantia:** a lei permite só uma modalidade (art. 37). Prefira caução, limitada a 3 aluguéis, ou seguro-fiança, em vez da fiança pessoal de dirigentes.
- **Benfeitorias:** necessárias e adaptações exigidas pelo Poder Público são indenizáveis, com direito de retenção (art. 35). O STJ admite cláusula de renúncia a essa indenização (Súmula 335), então resista se o locador propuser a renúncia.
- **Aprovações internas:**
  - Locação por mais de 5 anos exige autorização da Assembleia (Estatuto, art. 16, VI).
  - Locador ligado a dirigente exige aluguel comprovadamente de mercado, parecer do Conselho Fiscal e aprovação da Diretoria sem o interessado (Estatuto, art. 40).
- **Saída antecipada:** multa proporcional ao tempo restante (art. 4º), com cláusula opcional de dispensa em situações como interdição ou negativa de alvará.

## 14. Recibos de dízimos e ofertas

- **Por que o texto do recibo importa para a imunidade:** o recibo diz que a contribuição é voluntária, gratuita e sem contraprestação, e que não gera direito sobre o patrimônio nem restituição. Isso afasta a ideia de pagamento por serviço, protege a natureza religiosa das receitas e está em linha com o Estatuto (arts. 11 e 13, § 5º) e com a Declaração de Fé (seção VI).
- **Imposto de renda:** em regra, a pessoa física não pode deduzir doações a igrejas no IRPF, e a pessoa jurídica só deduz doações a entidades com qualificações específicas previstas em lei. Por isso o recibo diz que não é documento para dedução, "salvo nas hipóteses expressamente previstas em lei". Confirme com o contador antes de emitir recibos para empresas doadoras.
- **Controles:** numeração sequencial em duas vias, canhoto para lançamento contábil, recibo obrigatório acima de um valor definido e em doações designadas, recebimento só em contas no CNPJ da igreja e proibição de recibo com dados diferentes do recebido. Esses controles previnem fraude e uso do recibo para fins ilícitos.
- **Doações designadas:** quando o contribuinte indica uma finalidade (por exemplo, construção), a igreja deve aplicar o dinheiro nela. O recibo prevê o destino do saldo se a finalidade for concluída ou se tornar impossível, o que evita disputas futuras.
- **Devolução por erro material:** o dinheiro só volta em caso de erro comprovado (PIX em duplicidade, valor errado, transferência para o destinatário errado), com aprovação da Diretoria em ata e sempre para a conta de origem. Isso impede que "devoluções" sejam usadas para desviar recursos.
- **Doações de bens:** não usam este recibo, e sim termo de doação próprio (Regimento, art. 41, § 3º). Imóveis exigem escritura pública. Pela LC 227/2026, art. 149, a imunidade de ITCMD para doações a entidades religiosas alcança os bens ligados às finalidades essenciais. Verifique a lei do seu Estado.

## 15. Termo de doação de bens

**Base legal:** CC, arts. 538 a 564. Bens móveis podem ser doados por instrumento particular (art. 541). Imóveis de valor superior a 30 salários mínimos exigem escritura pública (art. 108), e a propriedade só passa para a igreja com o registro no Registro de Imóveis (art. 1.245). Por isso o termo cobre só bens móveis e traz, em anexo, um roteiro para a doação de imóveis. Os artigos do Código Civil foram citados sem conferência do texto, porque a base de legislação estava sem consultas.

**Proteções para a igreja:**
- **Contrapartidas:** a doação não traz contrapartida nem dá ao doador direito sobre o patrimônio ou a administração, conforme o Estatuto, art. 38, § 3º.
- **Encargos:** há duas opções. Na primeira, o doador apenas manifesta um desejo de uso, sem obrigar a igreja. Na segunda, há encargo formal (art. 553), mas a igreja pode vender o bem ou mudar o uso se ele ficar inservível ou desnecessário. Encargos desproporcionais ou que desvirtuem as finalidades da igreja são proibidos.
- **Reversão (art. 547):** o bem voltar ao doador se ele sobreviver à igreja fica apenas como opção, e recomendo não usar. Essa cláusula conflita com o destino do patrimônio previsto no Estatuto em caso de dissolução.
- **Declarações do doador:** a doação não pode comprometer a subsistência do doador (art. 548, que torna nula a doação de todos os bens sem reserva) nem ultrapassar a parte de que ele pode dispor em testamento (art. 549). Bens comuns do casal exigem a assinatura do cônjuge (art. 1.647, IV). O termo também traz declarações de propriedade, origem lícita e eventual vínculo com dirigentes.
- **Veículos:** placa, RENAVAM e chassi, a situação de débitos e a obrigação do doador de assinar a transferência no órgão de trânsito.

**ITCMD:** a doação à igreja é imune quando os bens se relacionam às finalidades essenciais (LC 227/2026, art. 149, I, "c", e § 1º, II, conferido anteriormente). O Estado pode exigir declaração ou procedimento próprio (§ 4º do mesmo artigo). No caso de imóveis, o reconhecimento da imunidade deve ser pedido antes da escritura.
