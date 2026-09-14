# Contexto do projeto CorelAI

Documento de continuidade criado em 13/09/2026, a partir da inspeção dos arquivos em `C:\CorelAI` e do histórico desta tarefa. O conteúdo descreve o estado encontrado nesta data; arquivos de entrada e previews podem ser regenerados posteriormente.

## 1. Objetivo e escopo

O projeto testa uma integração entre Codex e CorelDRAW mediada por arquivos locais. O usuário descreve uma alteração em linguagem natural, fornece uma representação visual da arte e uma descrição estruturada dos objetos, e Codex produz um plano textual que deverá ser interpretado por VBA no CorelDRAW.

O caso de uso atual é uma arte promocional de Antarctica 473 ml, com a campanha “leve MAIS por menos”, uma oferta “leve 12”, um preço principal, uma oferta secundária “leve 1” com outro preço e uma lata à direita.

O pedido de composição é transformar a arte em **500 × 100 cm**, equivalentes a **5000 × 1000 mm**, aproveitando o formato horizontal e preservando as proporções da lata e dos elementos de conteúdo. Não se devem inventar textos, preços ou produtos.

Esta tarefa de documentação cria somente este arquivo. Não executa comandos no CorelDRAW nem modifica os arquivos de entrada, layouts ou imagens existentes.

## 2. Arquitetura observada e componentes externos

Fluxo previsto:

```text
Documento no CorelDRAW
    -> exportação de cena.txt e preview.png
Pedido do usuário em pedido.txt
    -> leitura e análise pelo Codex
    -> plano em layout_vN.txt
    -> interpretação/aplicação por VBA no CorelDRAW
    -> nova exportação de cena e preview
    -> revisão visual e nova versão do plano
```

Responsabilidades:

- **CorelDRAW:** mantém os objetos reais, seus grupos, textos, curvas, bitmap, cores e geometria.
- **Exportador de cena:** fornece IDs e propriedades em texto. O arquivo atual também descreve filhos de grupos recursivamente.
- **Exportador de imagem:** fornece o aspecto visual da página para avaliar hierarquia, espaços e relações que a geometria sozinha não explica.
- **Codex:** interpreta o pedido, relaciona imagem e IDs e escreve o plano de transformação.
- **Interpretador VBA:** deve localizar os objetos e executar os comandos do plano. Seu código não foi encontrado nesta pasta.

A inspeção encontrou apenas arquivos de intercâmbio e instruções. Não foram encontrados arquivos `.cdr`, módulos `.bas`, macros `.gms`, código do parser, scripts de exportação, testes automatizados ou configuração de build. Não há evidência nesta pasta de um observador automático de arquivos, botão específico ou mecanismo de acionamento. A forma de disparar a macro deve ser confirmada com o usuário ou com o código externo.

O sucesso do teste de escrita local não prova, por si só, conexão direta com a API do CorelDRAW.

## 3. Inventário dos arquivos

| Arquivo | Papel e estado encontrado |
| --- | --- |
| `teste_corel.txt` | Teste inicial de escrita; contém `CONEXAO CODEX OK` seguido de término de linha no estado atual. |
| `comando.txt` | Saída do primeiro teste de classificação de pedido; contém `AJUSTAR`. É separado do protocolo de layout. |
| `pedido.txt` | Pedido de transformação para 500 × 100 cm, campanha em destaque e produto à direita. A frase termina em “organize as informações”, sem detalhamento adicional. |
| `cena.txt` | Cena atual: página de 5000 × 1000 mm, 17 objetos de nível superior e descrição de vários descendentes. Contém geometria, textos e propriedades adicionais. |
| `preview.png` | Imagem horizontal mais recente, visualmente compatível com a cena atual: campanha à esquerda, oferta central, preço principal grande e lata à direita. Não é mais o preview vertical original. |
| `preview_resultado.png` | Imagem de uma tentativa anterior: campanha dominante, preço principal pequeno e oferta secundária isolada. Não representa o mesmo estado de `cena.txt`. |
| `layout.txt` | Atualmente é um teste curto com `PAGE|500|700` e `SCALE|3.2|1.20`; não é o plano horizontal completo. |
| `layout_v2.txt` | Plano horizontal completo mais recente disponível, reproduzido adiante. Não existe `layout_v3.txt` na inspeção. |
| `INSTRUCOES_REVISAO.txt` | Regras de revisão, comandos permitidos, preservação de conteúdo e versionamento incremental. Sua versão atual proíbe sobrescrever versões existentes. |
| `CONTEXTO_PROJETO.md` | Este documento; registra fatos, decisões, limitações e próximos passos. |

Não se deve deduzir a versão de uma cena apenas pelo nome do arquivo. `preview.png` foi substituído ao longo do trabalho, enquanto `preview_resultado.png` permaneceu antigo.

## 4. Estado atual e divergências importantes

### 4.1 Cena e preview mais recentes

`cena.txt` declara `LARGURA_MM=5000.00`, `ALTURA_MM=1000.00` e `TOTAL_OBJETOS=17`. O preview atual tem composição horizontal de proporção 5:1.

A cena corresponde aos alvos principais do plano `layout_v2.txt`: campanha em X=1075, título em X=2780/Y=830, preço principal com 900 mm, lata com 960 mm de altura e fundo cobrindo 5000 × 1000 mm. Isso indica que o estado exportado incorpora essas transformações; não comprova qual macro ou sequência exata produziu o resultado.

### 4.2 Layout de teste atual

Conteúdo encontrado em `layout.txt`:

```text
PAGE|500|700
SCALE|3.2|1.20
```

Esse arquivo solicita uma página de **500 × 700 mm**, incompatível com o pedido horizontal de 5000 × 1000 mm. Também referencia um filho de grupo, `3.2`, e não o grupo principal `3`.

Não aplicar esse arquivo como se fosse o plano final. Pode ser um teste intencional do suporte a IDs hierárquicos, mas sua finalidade não está documentada. A cena disponível continua horizontal; a presença do comando não prova que o teste foi executado.

### 4.3 Mudança nas regras de versionamento

Em turnos anteriores, o usuário pediu explicitamente que `layout_v2.txt` fosse sobrescrito. Isso foi feito. A versão atual de `INSTRUCOES_REVISAO.txt` passou a determinar:

- nunca sobrescrever uma versão existente;
- usar a versão mais recente como ponto de partida;
- criar `layout_v2.txt` quando só houver `layout.txt`;
- criar `layout_v3.txt` quando `layout_v2.txt` já existir, e continuar a sequência.

Assim, **a próxima revisão deve gerar `layout_v3.txt`**, salvo nova instrução explícita do usuário. Não modificar `layout_v2.txt` automaticamente usando a regra antiga.

## 5. Modelo de dados de cena.txt

O arquivo usa cabeçalhos e campos `CHAVE=VALOR`, não JSON ou CSV. Os blocos de primeiro nível são `OBJETO_1`, `OBJETO_2` etc. Grupos podem conter `FILHOS`, blocos `FILHO_n` e `NIVEL`.

Campos observados:

- Página: `LARGURA_MM`, `ALTURA_MM`, `TOTAL_OBJETOS`.
- Identificação: `ID`, `ORDEM`, `TIPO`.
- Geometria: `X_MM`, `Y_MM`, `LARGURA_MM`, `ALTURA_MM`.
- Texto: `TEXTO`, `FONTE`, `TAMANHO_PT`.
- Aparência: `PREENCHIMENTO`, `PREENCHIMENTO_CMYK`, `CONTORNO`, `CONTORNO_MM`.
- Hierarquia: `FILHOS`, `FILHO_n`, `NIVEL`.

As coordenadas são compatíveis com centros dos objetos, em milímetros, com Y crescendo para cima: o fundo tem centro (2500, 500), enquanto o título aparece acima da faixa em Y=830. Essa leitura é sustentada pela correspondência entre cena e preview; o código exportador não foi inspecionado.

`ORDEM` é exportado, mas não foi confirmado se representa ordem de empilhamento, enumeração ou outro critério. Não reorganizar camadas com base apenas nesse campo.

IDs hierárquicos devem ser tratados como **strings**, por exemplo `3.2.2.1`. O ponto separa níveis e não é um separador decimal nesse contexto. Não converter IDs para números reais nem pressupor que sejam identificadores persistentes do CorelDRAW entre exportações.

O histórico começou com números decimais usando vírgula; a exportação atual usa ponto nas medidas. `PREENCHIMENTO_CMYK`, por sua vez, usa vírgulas para separar componentes. Um parser precisa distinguir esses contextos.

## 6. Mapeamento visual dos objetos

| IDs | Elemento | Orientação para transformação |
| --- | --- | --- |
| `1,2` | Painel da campanha “leve MAIS por menos” | Mover/escalar juntos para preservar relação entre painel e conteúdo. Ambos têm os mesmos limites externos. |
| `2.1` | Texto “por menos” | Filho do grupo 2. |
| `2.2` | Texto “leve MAIS” | Filho do grupo 2; o preview mostra cores diferentes dentro do texto. |
| `2.3` | Retângulo interno da campanha | Filho do grupo 2; não remover por parecer duplicado do ID 1. |
| `3` | Conjunto completo da oferta “leve 1” | Preservar como conjunto, salvo pedido explícito de ajuste interno. |
| `3.1` | Texto “leve 1” | Associado à faixa 3.3 e ao selo 3.2. |
| `3.2` | Selo de preço secundário | Grupo com oito filhos diretos, incluindo centavos em outro subgrupo. |
| `3.3` | Faixa vermelha da oferta secundária | Geometria que pode admitir STRETCH, mantendo o conteúdo dentro. |
| `4,6` | Algarismos grandes do preço principal | Curvas; não há texto editável com o valor nesses blocos. |
| `5` | Centavos do preço principal | Grupo com filhos 5.1 e 5.2. |
| `7` | Vírgula do preço principal | Manter junto aos demais componentes do preço. |
| `8` | “CADA” do preço principal | Identificação visual; armazenado como curva. |
| `9` | “R$” do preço principal | Identificação visual; armazenado como curva. |
| `10,11` | Formas circulares do selo principal | Preservar junto de 4–9; o ID 10 tem 900 × 900 mm. |
| `12` | “ANTARCTICA 473ml” | Texto, centralizado sobre a faixa principal. |
| `13,14` | “leve” e “12” | Ajustar juntos dentro do retângulo 16. |
| `15` | Lata Antarctica | Bitmap; escala proporcional obrigatória. |
| `16` | Faixa vermelha principal | Pode mudar largura e altura via STRETCH. |
| `17` | Fundo amarelo | Deve preencher a página com FILLPAGE. |

Os preços aparecem visualmente como algarismos “88,88”. Não há autorização para substituir esse conteúdo por valores comerciais. A cena descreve curvas, portanto a leitura do preço não deve ser inferida de propriedades textuais inexistentes.

## 7. Protocolo de comandos

O plano usa uma instrução por linha, campos separados por `|`, medidas em milímetros e fatores adimensionais. Os exemplos atuais usam ponto decimal e listas de IDs separadas por vírgula. Evitar comentários ou comandos extras no arquivo consumido pelo VBA, pois sua aceitação não foi verificada.

Os comandos abaixo são os **declarados pelas instruções locais**. A implementação e os detalhes do parser não estão disponíveis para auditoria.

| Sintaxe | Finalidade esperada e cuidados |
| --- | --- |
| `PAGE|largura_mm|altura_mm` | Definir dimensões da página. Para o pedido principal: `PAGE|5000|1000`. |
| `FILLPAGE|ID` | Ajustar o objeto de fundo para cobrir a página. Usado com 17. |
| `OBJECT|ID|CENTER_X_MM|CENTER_Y_MM|WIDTH_MM` | Posicionar e dimensionar proporcionalmente um objeto pela largura. A altura acompanha a proporção corrente. Pode apontar para um grupo inteiro. |
| `SET|ID1,ID2,...|CENTER_X_MM|CENTER_Y_MM|WIDTH_MM` | Transformar um conjunto mantendo a composição relativa. A largura é a do conjunto, não a de cada membro. Não implica necessariamente agrupamento permanente no documento. |
| `STRETCH|ID|CENTER_X_MM|CENTER_Y_MM|WIDTH_MM|HEIGHT_MM` | Definir largura e altura independentemente. Permitido para fundos, faixas e geometria apropriada; não deformar produtos, logos ou texto. |
| `FITINSIDE|ID1,ID2,...|ID_CONTAINER|FATOR` | Ajustar conteúdo relacionado dentro de um contêiner. Espera-se escala proporcional do conjunto e posicionamento contido. A definição exata do fator e centralização deve ser confirmada no VBA. |
| `SCALE|ID|FATOR` | Ajuste relativo de escala. Reservar para pequenos ajustes de hierarquia. A âncora de escala precisa ser confirmada no VBA. |

`AJUSTAR`, em `comando.txt`, pertence ao teste inicial de classificação e não integra esta lista de instruções de layout.

### Ordem e reexecução

Tratar as linhas como uma sequência: definir página, ajustar fundo, posicionar elementos e preparar um contêiner antes de executar FITINSIDE nele. Não ampliar um texto isolado depois de FITINSIDE sem verificar novamente a contenção.

SCALE é relativo: reexecutar `SCALE|3.2|1.20` pode ampliar novamente o objeto a cada execução. Mesmo os comandos com dimensões absolutas dependem da geometria e da hierarquia correntes. Não assumir que todo plano seja idempotente.

Não selecionar simultaneamente um grupo pai e um descendente na mesma transformação sem conhecer o tratamento do VBA: pode ocorrer transformação duplicada. O conjunto `1,2` usa dois objetos de primeiro nível; `2.1`, por exemplo, é descendente de 2.

## 8. Plano completo mais recente disponível

Conteúdo de `layout_v2.txt` no momento da documentação:

```text
PAGE|5000|1000
FILLPAGE|17
SET|1,2|1075|500|1950
OBJECT|12|2780|830|1100
STRETCH|16|2780|640|1250|220
FITINSIDE|13,14|16|0.82
SET|4,5,6,7,8,9,10,11|3900|500|900
OBJECT|3|2780|265|900
OBJECT|15|4740|500|422.54
```

Leitura da composição atual:

- Campanha à esquerda, centro (1075, 500), largura 1950 e altura exportada 643,76 mm.
- Nome do produto centralizado em X=2780, acima da faixa; dimensões 1100 × 90,47 mm.
- Faixa em (2780, 640), com 1250 × 220 mm; “leve 12” contido nela.
- Preço principal à direita da oferta, centro (3900, 500), diâmetro 900 mm.
- Oferta secundária abaixo do título/faixa, centro (2780, 265), largura 900 mm; altura atual exportada 432,48 mm.
- Lata ancorada à direita em (4740, 500), com 422,54 × 960 mm.
- Fundo em (2500, 500), cobrindo 5000 × 1000 mm.

Relações geométricas úteis, calculadas a partir da cena atual:

- Preço principal ocupa Y=50 a 950 mm.
- Lata ocupa Y=20 a 980 mm, com margem direita aproximada de 48,73 mm.
- Faixa principal ocupa Y=530 a 750 mm.
- Título começa aproximadamente em Y=784,77 mm: há cerca de 34,77 mm entre faixa e título.
- Oferta secundária termina em Y=481,24 mm: há cerca de 48,76 mm até a faixa principal.
- Faixa termina em X=3405 mm e selo principal começa em X=3450 mm: intervalo de 45 mm.
- Os limites geométricos principais não se sobrepõem, mas limites não substituem avaliação óptica de letras, contornos e transparências.

## 9. Decisões tomadas ao longo do trabalho

1. Testar a escrita local com `teste_corel.txt`.
2. Interpretar um pedido inicial de aumento com margem de 10 mm como `AJUSTAR`. Essa margem era daquele teste; não é requisito explícito do pedido horizontal atual.
3. Trocar a abordagem de um comando genérico por um plano por IDs, com posições e larguras explícitas.
4. Interpretar 500 × 100 cm corretamente como 5000 × 1000 mm.
5. Tratar campanha e componentes de preço como conjuntos para evitar desmontagem visual.
6. Preservar a proporção da lata e conteúdo; permitir deformação controlada somente de fundos e faixas conforme as instruções posteriores.
7. Substituir o fundo proporcional excedente usado no primeiro plano por FILLPAGE, adequado ao fundo geométrico.
8. Ampliar o selo principal, inicialmente de 470 mm, para 900 mm, reduzindo a dominância relativa da campanha.
9. Relacionar título e “leve 12” por proximidade e eixo central comum; alinhar a oferta secundária abaixo.
10. Usar STRETCH e FITINSIDE para eliminar o excesso de faixa vazia e conter os textos. Remover do plano revisado a ampliação isolada de “leve” após o ajuste conjunto.
11. Aumentar a lata para 960 mm de altura, funcionando como âncora visual à direita.
12. Adotar, pelas instruções atuais, versionamento incremental para revisões futuras.

## 10. Limitações e riscos conhecidos

- **Previews fora de sincronia:** `preview_resultado.png` não corresponde à cena mais recente. Não avaliar o estado atual exclusivamente por esse arquivo.
- **Original vertical não preservado no arquivo atual:** no início, `preview.png` mostrava uma página vertical; agora ele contém uma imagem horizontal. O original não foi encontrado como outro arquivo local.
- **Código externo ausente:** não é possível confirmar suporte efetivo a todos os comandos, transações, undo, tratamento de erros, IDs hierárquicos ou seleção de documento/página.
- **IDs sem garantia de estabilidade:** a estrutura sugere caminhos de grupos. Reordenar ou reagrupar objetos pode alterar o mapeamento; exportar novamente antes de novas operações.
- **Mudanças internas alteram proporções de grupos:** o grupo 3 tinha originalmente 278,04 × 115,98 mm. Atualmente tem 900 × 432,48 mm. Não reutilizar sua proporção histórica para calcular novos limites. Há um teste de SCALE do filho 3.2, mas a causa exata dessa diferença não foi comprovada.
- **Cena não é uma descrição visual completa:** campos de preenchimento de grupo podem não representar cada filho ou cor de trechos de texto. O preview mostra “MAIS” amarelo embora o campo agregado do texto não descreva essa diferença.
- **Tamanhos de fonte negativos:** a cena exporta valores negativos para alguns textos, inclusive campanha e “12”. O significado não foi confirmado. Não corrigir tipografia usando esses números sem inspecionar o exportador/API.
- **Codificação:** a leitura padrão de `pedido.txt` apresentou caracteres de substituição em palavras acentuadas. Determinar a codificação real antes de implementar leitura/escrita no VBA. Não converter arquivos existentes automaticamente.
- **Separadores numéricos:** a cena antiga usava vírgula e a atual usa ponto. Evitar dependência silenciosa das configurações regionais do Windows.
- **Fonte e bitmap:** não foi verificada disponibilidade da fonte Futura Md BT, resolução efetiva da lata em tamanho de impressão, perfis de cor ou qualidade de saída a 5 metros de largura.
- **Contornos e escala:** não está comprovado se o VBA escala contornos junto com objetos. Isso pode explicar diferenças de espessura visual entre versões.
- **Sem validação de impressão:** margens são decisões visuais; não há especificação de sangria, acabamento ou área segura de produção.
- **Sem confirmação de executor automático:** gravar um layout não equivale a aplicá-lo. Não afirmar execução apenas por ter criado o arquivo.

## 11. Procedimento recomendado para a próxima revisão

1. Reler `INSTRUCOES_REVISAO.txt`, `pedido.txt` e o inventário, pois eles mudaram durante o trabalho.
2. Confirmar que cena e preview foram exportados juntos do mesmo documento e página.
3. Usar `layout_v2.txt` como plano completo mais recente; não usar o `layout.txt` de teste como especificação do banner final.
4. Mapear os objetos com base na cena atual, incluindo filhos se a alteração exigir ajustes internos.
5. Avaliar hierarquia visual, relação título/faixa/preço, distinção entre preços, margens e presença da lata.
6. Criar `layout_v3.txt` ou o próximo número livre, preservando as versões anteriores.
7. Validar IDs, aridade dos comandos, dimensões positivas, unidades, limites e contenção dos textos antes da aplicação.
8. Aplicar no CorelDRAW somente no escopo autorizado, com o executor confirmado. Nesta documentação nenhuma aplicação é solicitada.
9. Exportar novamente cena e imagem; conferir visualmente o resultado e comparar com o plano usado.

## 12. Próximos passos técnicos sugeridos

Estas ações são propostas, não implementações já realizadas:

1. Localizar e documentar o código VBA do exportador e do interpretador, incluindo como são acionados.
2. Confirmar a semântica exata de FITINSIDE, FILLPAGE e SCALE, além da resolução de IDs como `3.2` e `3.2.2.1`.
3. Esclarecer o objetivo do teste `PAGE|500|700` antes de executá-lo novamente.
4. Vincular cada exportação ao layout aplicado, com identificador de revisão e data, evitando confundir previews antigos e novos.
5. Padronizar a codificação e o parser numérico, preservando IDs como strings.
6. Adicionar validação prévia de comandos e mensagens de erro para IDs inexistentes, argumentos inválidos e contêineres inadequados.
7. Verificar a possibilidade de agrupar a execução em uma operação de undo e impedir execução parcial após erro.
8. Documentar as convenções de exportação de fontes, cores, contornos e grupos, especialmente os tamanhos negativos.
9. Validar, antes de produção, a resolução do bitmap, fontes, cores e requisitos da gráfica.

## 13. Orientação rápida para continuidade

O trabalho atual é um fluxo por arquivos, não uma integração de código presente nesta pasta. A cena já está em 5000 × 1000 mm e o preview atual mostra o preço principal grande. `layout_v2.txt` guarda o plano completo; `layout.txt` foi substituído por um teste de página vertical e escala de filho de grupo. A próxima revisão deve criar uma nova versão, começando por `layout_v3.txt`, e deve usar os dados atuais, sem reconstruir a cena a partir das dimensões antigas do histórico.
