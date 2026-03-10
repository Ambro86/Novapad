# Changelog

Versão 0.6.8 – 2026-03-04

Novidades
• Adicionado um novo item no menu Reproduzir que permite transcrever qualquer ficheiro de áudio ou vídeo com o Whisper. Nas Opções existe uma nova secção chamada «IA e Transcrição», onde é possível escolher o modelo, ativar o suporte opcional a CUDA para placas gráficas NVIDIA, manter o idioma original e ativar ou desativar as marcas temporais.

Melhorias
• Adicionada a gestão de perfis de voz em Opções > Voz: é possível adicionar, renomear e eliminar perfis.
• Alargadas em Opções > Áudio as opções do intervalo de retrocesso durante a reprodução, com novos valores de 1 segundo até 2 horas.
• Adicionada a tradução russa graças a Dmitriy.
• Adicionada em Opções > Áudio uma nova opção para escolher o formato do nome das partes do audiolivro: `Título + número`, `Somente número` ou `Número + título`.
• Adicionada no menu de contexto dos artigos RSS a ação para adicionar o artigo aos favoritos.
• A fonte RSS "Favoritos" pode ser eliminada e é recriada automaticamente quando um novo artigo é adicionado aos favoritos.
• Adicionados atalhos de teclado RSS para mover as fontes para cima/para baixo: `Ctrl+Shift+Seta para cima` e `Ctrl+Shift+Seta para baixo`.
• Melhorada a janela RSS com uma pré-visualização integrada do artigo, permitindo consultar o texto diretamente ali e alcançá-lo rapidamente com Tab antes de abrir o artigo completo no editor.
Correções
• Corrigido um erro em que a lista de ficheiros abertos era mostrada no menu Ajuda em vez do menu Janela.
• Corrigido um caso limite no streaming em que a reprodução podia iniciar, mas a janela “Transferência de streaming” permanecia aberta quando o ficheiro descarregado já correspondia ao formato de destino.
• Corrigido o comportamento de conversão no streaming MP3: quando o stream já está em MP3 e o utilizador escolhe um bitrate MP3 explícito (por exemplo 128 kbps), o Sonarpad agora recodifica para o bitrate selecionado em vez de saltar a conversão.
• Corrigido o atalho `Alt+Shift+L`: agora abre corretamente a lista de capítulos durante a reprodução.
• Adicionado suporte a capítulos de podcast incorporados em ficheiros multimédia locais (por exemplo, metadados de capítulos MP3): quando o feed/URL não fornece capítulos, o Sonarpad passa a carregá-los do ficheiro descarregado em segundo plano, permitindo início imediato da reprodução e aplicação dos capítulos assim que ficam disponíveis.
• Corrigido o carregamento de capítulos para episódios de podcast descarregados e abertos como ficheiros multimédia locais normais: os capítulos incorporados passam agora a estar disponíveis também nesse caso, e não apenas quando a reprodução começa a partir da janela Podcasts.
• Corrigida a finalização dos audiolivros MP3 com SAPI4 e SAPI5: o ficheiro final passa agora a ser finalizado corretamente, evitando ficheiros incompletos ou frágeis após exportações longas.
• Adicionada uma barra de progresso explícita para a fase de finalização em todos os modos de criação de audiolivros: após a criação, o Sonarpad anuncia e mostra a finalização com progresso visível.
• Corrigido um erro nas vozes de diálogo: os parâmetros de velocidade/tom/volume da primeira e da segunda voz de diálogo são agora aplicados corretamente durante a síntese.
• Melhorada a deteção de codificação para ficheiros japoneses `.txt`: adicionado fallback seguro Shift_JIS/CP932 em casos de mojibake, preservando o comportamento existente para UTF/diacríticos/chinês.
• Refatoração interna de segurança: conversão para implementações safe sempre que possível e redução drástica das linhas de código unsafe.
Versão 0.6.7 – 2026-03-02
Melhorias
• Ora il programma riesce a gestire Sostituisci tutto in modo massivo su file grandi con un gran numero di sostituzioni.
• Tradução polaca atualizada graças ao DJ Graco.
• Adicionada a tradução lituana.
• Adicionada a tradução chinesa.
• A partir de agora, compilações beta frequentes serão publicadas na seção Releases do projeto, para que os usuários possam testar as novas alterações antes da próxima versão estável.
• Adicionado o atalho `Ctrl+.` para inserir o caractere de reticências (…).
• Melhorado o suporte a capítulos de podcast: a navegação por capítulos está agora mais fiável, incluindo episódios diretos/streaming em que os capítulos não estão incorporados no ficheiro MP3, usando metadados de capítulos do feed/URL como fallback quando disponíveis. Adicionados os atalhos `Ctrl+Alt+PageUp` (capítulo anterior) e `Ctrl+Alt+PageDown` (capítulo seguinte).
• Reorganizadas as pastas de saída em `Documentos\\Sonarpad`: os ficheiros agora são guardados nas subpastas dedicadas `audiobooks`, `documents`, `recordings` e `media`, com migração automática dos caminhos antigos.
• Melhorado o suporte para ficheiros de texto muito grandes (incluindo 60 MB): abertura e navegação linha a linha mais fluídas, especialmente com leitores de ecrã.
• Guias atualizados para todos os idiomas e recursos de localização atualizados em toda a app, incluindo textos de doações e traduções do instalador NSIS (novas strings em chinês simplificado e lituano, além da conclusão da tradução ucraniana do setup).
• Adicionado suporte global de proxy de rede (HTTP/HTTPS e SOCKS5/SOCKS5H) para funcionalidades online, com validação ao guardar Opções: proxies inválidos são avisados e removidos automaticamente.
• Adicionada uma nova função em Ferramentas: "Reproduzir áudio por streaming...", que permite colar um URL (YouTube ou link multimédia direto), escolher o formato de saída e o perfil de qualidade/bitrate (incluindo qualidade/bitrate original para MP3 e MP4) e iniciar a reprodução no leitor de áudio do Sonarpad.
• Adicionado suporte à tecla multimédia de sistema Reproduzir/Pausar (auscultadores/teclado): agora controla tanto a reprodução multimédia como a pausa/retoma da leitura de texto (com prioridade para o leitor multimédia quando ambos estão ativos).
• Adicionada uma nova opção em Ficheiro > Ficheiros recentes: "Limpar ficheiros recentes" para esvaziar rapidamente a lista de documentos recentes.
• Ampliadas as opções de bitrate na conversão de áudio e na gravação de podcast: adicionados valores mais baixos (64/96 kbps) e MP3 estendido até 320 kbps, com validação e tratamento do encoder alinhados.
• Ampliadas as opções de divisão de audiolivros por tempo até 60 minutos.
• Melhorada a divisão de audiolivros por partes: agora o número de partes pode ser inserido manualmente, com validação de 1 a 100.
• Adicionado o novo modo Ver > Somente leitura para bloquear edições acidentais no texto, mantendo leitura e navegação completas dos documentos.
• Adicionada uma barra de progresso acessível durante as atualizações do programa, para que leitores de ecrã possam acompanhar em tempo real o progresso da transferência.
• Adicionada uma nova barra de estado discreta na janela principal com contagem de caracteres, palavras e posição linha/coluna (por exemplo: "Caracteres (com espaços): 11. | Palavras: 2. | Ln 1, Col 12"), sem interferir com o foco do NVDA.
• Adicionada uma nova opção no menu Ver para quebra automática de linha, permitindo ativar/desativar rapidamente sem abrir as Opções.
• Adicionadas em Editar > Texto novas ações para aumentar/reduzir recuo, com atalhos Ctrl+Shift+. (indentar) e Ctrl+Shift+, (desindentar), porque quando “Mostrar vozes no editor” está ativo a tecla Tab fica reservada para a navegação do painel de vozes.
• Adicionada a exibição localizada de data e hora em artigos RSS e episódios de podcast, com formato adaptado ao idioma da interface.
• Adicionada no menu de contexto RSS uma nova ação para partilhar por e-mail o artigo selecionado.
• Adicionadas opções granulares de confirmação de remoção em Opções > RSS e podcast: para RSS (feed/artigo/ambos/nenhum) e para Podcasts (podcast/episódio/ambos/nenhum).
• Adicionada cópia rápida de RSS configurável com Ctrl+C (Opções > RSS e podcast): copiar título, URL, conteúdo do artigo ou tudo junto.
• Fluxo de RSS unificado: “Adicionar fonte” agora aceita tanto URL de feed quanto palavras-chave (com geração automática do feed do Google News), sem necessidade de pesquisa separada.
• Ao premir Ctrl+A, o programa agora anuncia a conclusão da ação para um feedback mais claro em leitores de ecrã.
• Adicionado Shift+F3 para "Localizar anterior" no menu Editar, em complemento ao F3 "Localizar seguinte".
• Melhorada a mensagem de substituição com singular/plural corretos (por exemplo, “1 substituição realizada” vs “2 substituições realizadas”).
• Adicionada na janela do dicionário a seleção de idioma de pesquisa, com Auto (idioma da interface) por padrão e possibilidade de escolha manual.
• Adicionada uma nova aba de Atalhos nas Opções para personalizar combinações de teclas, com deteção de conflitos e aviso quando um atalho já está atribuído a outra ação.
• Adicionado suporte inicial a parâmetros de linha de comandos: `-h`/`--help` mostram a ajuda rápida e `--version` mostra a versão do programa.
• Melhorada a clareza do ajuste manual de velocidade e tom: os campos manuais agora usam uma escala centrada em 100, onde 100 corresponde ao valor normal.
• Melhorada a seleção de vozes Microsoft em Opções > Voz e no painel de vozes do editor: foi adicionada uma combobox de idioma localizada para filtrar vozes por idioma, mantendo o modo “apenas vozes multilíngues” como lista única sem divisão por idioma (com a combobox de idioma oculta quando ativa).
• Adicionada a configuração de voz para diálogos em Opções > Voz com navegação completa por Tab, usando o mesmo modelo de vozes da interface principal (motor, filtro de idioma Edge, voz e velocidade/tom/volume com etiquetas); adicionada também uma segunda voz de diálogos opcional com os mesmos controles (motor, filtro de idioma Edge, voz, velocidade/tom/volume) para alternar diálogos; as regras de diálogos são guardadas na configuração `.ini`, sem modificar o texto do documento.
• Melhorada a etiqueta de Desfazer: a opção Editar > Desfazer agora mostra qual ação será desfeita (por exemplo, edição de texto, comentar/descomentar linhas ou inserção de tag de voz), mantendo-se indisponível quando não há nada para desfazer.
Correções de bugs
• Corrigido o suporte de abertura RTF: os ficheiros `.rtf` agora são extraídos e mostrados como texto legível, em vez de markup RTF bruto (ex.: `{\\rtf1...}`).
• Corrigida a abertura de ficheiros de texto chineses em codificação GB18030/GBK: o Sonarpad agora deteta e descodifica corretamente, evitando texto ilegível (mojibake).
• Melhorada a criação de audiolivros M4B com metadados e marcadores de capítulos; corrigido o problema "chipmunk" (voz demasiado aguda/rápida) nos ficheiros M4B gerados.
• Corrigida a interface de bitrate na janela de gravação de audiolivro: removidos textos hardcoded em italiano e adicionada a opção 64 kbps entre os bitrates selecionáveis.
• Corrigido "Guardar tudo" (Ctrl+Shift+S): agora todos os documentos abertos modificados são detetados de forma fiável (incluindo separadores novos/sem guardar) e o Guardar tudo grava cada um corretamente, abrindo "Guardar como" quando necessário.
• Corrigida a ordenação dos artigos RSS do Google News: quando a data está disponível, os artigos agora são mostrados do mais recente para o mais antigo.
• Corrigida a associação de etiquetas no NVDA na janela do dicionário: o campo de pesquisa e a lista de idioma agora anunciam a etiqueta correta.
• Corrigida a navegação por teclado na janela Propriedades de RSS/Podcast: Tab/Shift+Tab agora alcançam o botão OK, Enter ativa o OK, Esc fecha com segurança e o foco volta corretamente à lista RSS/Podcast.
• Corrigido o histórico de desfazer em RSS/Podcast: o Ctrl+Z agora suporta desfazer em múltiplos níveis para remoções (artigos/episódios e fontes), e não apenas a última ação.
• Melhorados os anúncios de remoção em RSS/Podcast com mensagens explícitas (RSS removido, artigo RSS removido, episódio de podcast removido).
• Melhorado o comportamento de foco após remover/desfazer em RSS/Podcast: no RSS, o primeiro feed volta a ser selecionado de forma fiável quando necessário, e foram reduzidas repetições de anúncios do leitor de ecrã durante a re-seleção atrasada.

Versão 0.6.6 – 2026-02-13
Melhorias
• Adicionada "Formatação automática para TTS" no menu Editar para preparar rapidamente o texto para voz (remove markdown/aspas e recompõe linhas quebradas).
• Melhorada a inserção de tags de voz: quando há texto selecionado, as tags passam a ser aplicadas corretamente tanto em uma única linha quanto em seleção multilinha.
• Adicionada uma opção nas Configurações de áudio para escolher a pasta padrão de gravação de audiolivros (padrão: Documentos\\Sonarpad Audiobooks).
• Na janela de gravação de audiolivro, quando a divisão em partes está ativa, foi adicionada uma nova opção (ativada por padrão) para criar uma subpasta dedicada às partes geradas.
• A exportação de audiolivros agora guarda MP3 em estéreo com bitrate escolhido pelo utilizador para vozes Edge, SAPI5 e SAPI4.
• Adicionado suporte a vozes SAPI5 de 32 bits via bridge, para usar também vozes disponíveis apenas em motores de 32 bits.
• Reorganizadas as funcionalidades de voz num menu dedicado "Voz e áudio" e adicionada/esclarecida a opção "Converter áudio", útil para converter qualquer ficheiro multimédia suportado para MP3, AAC, OGG, Opus, FLAC, WAV e AIFF.
• Adicionada a remoção de artigos RSS individuais e episódios de podcast individuais (tecla Delete + menu de contexto com confirmação), sem remover toda a fonte RSS/podcast, com anulação da última remoção (artigo/episódio individual ou fonte RSS/podcast completa).
• Adicionada a exportação de feeds RSS para OPML na janela RSS, para guardar e reimportar facilmente as fontes atuais.
• Adicionada a função "Pesquisar RSS por palavra-chave" na janela RSS: ao inserir uma palavra-chave, o Sonarpad gera automaticamente o URL RSS do Google News e abre a janela de adicionar fonte já pré-preenchida, permitindo criar um feed temático num único passo.
• Adicionada a tradução sérvia graças a Mila Kuran.
• Adicionada a tradução ucraniana graças a Ivan Shtefuriak.
• Adicionada a abertura múltipla de ficheiros multimédia: ao abrir vários ficheiros de uma vez é criada uma fila de reprodução em vez de substituir o ficheiro atual.
• Adicionados atalhos de avanço/retrocesso variável durante a reprodução: com base de 1 minuto, Esquerda/Direita avança 60s, Shift+Esquerda/Direita avança 20s e Ctrl+Esquerda/Direita avança 3 minutos.
• Adicionados atalhos de faixa anterior/seguinte no leitor: Ctrl+PageUp e Ctrl+PageDown.
• Adicionada a opção "Repor volume" e agrupadas as ações de reposição num submenu dedicado "Repor" em Reprodução, juntamente com "Repor velocidade" e "Repor tom".
• Melhorias no instalador: o setup.exe agora permite escolher entre associar todos os tipos de ficheiro suportados ou selecionar manualmente as extensões; o MSI também passa a oferecer seleção por extensão na árvore de funcionalidades (o padrão mantém-se: tudo ativo).
• Adicionado o novo menu "Janela" com a opção "Documentos abertos..." para alternar rapidamente para qualquer ficheiro atualmente aberto.
• Atualizada a opção Ver > Fonte: o seletor completo foi substituído por um submenu rápido com fontes comuns (Arial, Calibri, Consolas, Segoe UI, Tahoma, Verdana, Times New Roman, Georgia), mantendo o tamanho de texto atual.
• Melhorada a leitura de RSS e podcasts com dois avisos distintos: os nós da fonte anunciam "novos itens" quando um feed/podcast tem novidades, enquanto artigos RSS e episódios de podcast individuais anunciam "não lido"/"não reproduzido"; este comportamento pode ser desativado nas Opções.
Correções de bugs
• Corrigida a extração de texto EPUB para livros com comentários HTML inline (<!-- ... -->): o texto dos capítulos agora é analisado corretamente em vez de ser parcialmente ou totalmente ignorado.
• Corrigido o dicionário Wiktionary em espanhol e o cache do dicionário: palavras como "agua" agora são encontradas corretamente e entradas antigas de "Palavra não encontrada" não são mais reutilizadas.
• Corrigida a codificação na importação de artigos RSS para algumas fontes em espanhol (ex.: El Mundo): acentos e "ñ" agora são preservados corretamente no editor temporário.
• Corrigida a descodificação ANSI de ficheiros da Europa Central (ex.: checo/polaco): o Sonarpad agora distingue melhor UTF-8 e ANSI e escolhe a code page correta (incluindo Windows-1250), evitando diacríticos corrompidos.
• Corrigida a persistência de fontes RSS com parâmetros na URL (ex.: `rss.aspx?c=...`): esses feeds agora são guardados e restaurados corretamente após reiniciar o Sonarpad.
• Corrigida a abertura de ficheiros ponteiro do Google Drive (`.gdoc`, `.gsheet`, `.gslides`) a partir do menu de contexto do Explorador: quando a leitura direta falha com “Incorrect function (os error 1)”, o Sonarpad agora usa fallback por shell-open e o documento abre corretamente.
• Corrigida a leitura de ficheiros Excel legacy `.xls` (Excel 2010): ficheiros binários antigos são agora detetados/descodificados corretamente em vez de mostrar texto corrompido (ex.: `ÐÏ_à¡±...`).
• Corrigido o fluxo de anúncio do corretor ortográfico: os erros voltam a ser anunciados ao rever o texto mais tarde, e o mesmo erro é novamente reportado se for apagado e reescrito.
• Corrigidas as operações de texto por linha (ex.: Ctrl+Q / Ctrl+Shift+Q, ordenar/inverter/linhas únicas/unir linhas): ao selecionar apenas uma linha com Shift+Seta para baixo, as linhas adjacentes não são mais unidas nem truncadas.
• Corrigido o comportamento em seleções multilinha nas operações por linha (Ctrl+Q / Ctrl+Shift+Q e ferramentas relacionadas): quando o RichEdit devolve separadores de linha apenas com CR, agora são normalizados corretamente e todas as linhas selecionadas são processadas sem cortar o primeiro carácter.
• Ampliada a normalização de entrada TTS para símbolos visíveis de espaço/tab/nova linha (␠/U+2420, ␣/U+2423, ␉/U+2409, ␊/U+240A, ␍/U+240D, ␤/U+2424), que com vozes multilíngues podiam causar repetição de parágrafos.
• Refinada a sanitização do texto para Edge TTS com uma única pipeline de validação: normalização de espaços estranhos/invisíveis, compactação de sequências longas de pontuação (como "...", "!!!", "???") e descarte de trechos compostos só por pontuação para evitar loops de reprodução.
• Corrigido o anúncio do tempo de reprodução (Ctrl+I) para streams MP3/podcast: o tempo atual agora é limitado à duração da faixa, e a reprodução é interrompida automaticamente se a posição ultrapassar o fim.
• Melhorada a cobertura de localização do instalador: o setup.exe agora inclui também checo, polaco, francês e sérvio, enquanto o MSI permanece como um único pacote en-US para evitar confusão nas releases.
• Corrigida a limpeza na desinstalação das entradas do menu de contexto: "Abrir com Sonarpad" agora é removido de forma confiável, inclusive em cenários legados de registo.
• Corrigida a fiabilidade de pausar/retomar no SAPI5: a pausa com F4 agora funciona corretamente e, ao retomar, volta ao ponto esperado em vez de reiniciar do início.
• Corrigido o fluxo pausar + procurar + retomar na reprodução multimédia: após pausar e avançar/recuar com Esquerda/Direita, ao premir Espaço a reprodução retoma de forma fiável na posição atual em vez de parar ou reiniciar do início.

Versão 0.6.5 – 2026-02-07
Melhorias
• Tradução em espanhol aprimorada graças a Arturo Fernandez Rivas.
• Adicionada uma opção para dividir audiolivros EPUB por capítulos.
• As importações RSS agora usam uma aba temporária dedicada (título localizado); Salvar como a converte em um documento normal.
• As mensagens do leitor de ecrã agora também são enviadas ao JAWS quando disponível.
Correções de bugs
• A leitura a partir do cursor (F5) agora começa exatamente no cursor. Antes podia começar algumas linhas acima porque o deslocamento do cursor não correspondia às posições CRLF/UTF-16.
• Corrigido um problema de redesenho: ao digitar sobre uma seleção, o texto anterior podia desaparecer até mover a seleção.
• Corrigido o parser de capítulos EPUB: páginas de capa ou apenas com imagens não geram mais leitura de CSS (ex.: "padding") nem títulos "Sconosciuto".
• Corrigida a falha ao dividir por tempo audiolivros a partir de EPUB: o Edge TTS podia falhar com chunks vazios ou muito longos ("Edge audio not sent").
• Os artigos RSS agora decodificam entidades HTML (por ex. &quot;, &amp;, &lt;, &gt;).
• Salvar/Salvar como agora sugere o nome do arquivo existente ao salvar formatos que não devem ser sobrescritos (ex.: EPUB), em vez da primeira linha.
• Corrigido um problema em que podcasts com novos episódios não eram anunciados como não reproduzidos, e renomeado "Não ouvido" para "Não reproduzido" por ser mais profissional.

Versão 0.6.4 – 2026-02-05
Melhorias
• O programa foi renomeado para Sonarpad para dar maior ênfase ao som e ao áudio, que são a chave deste programa.
• Adicionada a seleção de faixas de áudio no menu Reprodução para arquivos multimídia com múltiplas faixas de áudio (ex. MKV com vários idiomas).
• Os podcasts agora indicam claramente os não ouvidos com o prefixo "Não ouvido" antes do nome.
• Novo sistema de tags para mudar a voz no texto. Exemplos:
  - Vozes Microsoft (Edge): <voice edge it-IT-IsabellaNeural>Olá</voice>
  - Vozes SAPI5: <voice sapi5 Microsoft Helena Desktop>Olá</voice>
  - Vozes SAPI4: <voice sapi4 #1>Olá</voice>
  - Com velocidade/tom/volume: <voice edge it-IT-ElsaNeural speed=-20 pitch=-5 volume=-10>Olá</voice>
• Categorias de podcasts enriquecidas.
• Adicionada uma opção no menu de contexto para criar um audiolivro a partir da seleção.
• Adicionada a divisão de audiolivros por duração, com a possibilidade de escolher o nome do primeiro arquivo.
• Rótulo do autor localizado na leitura de artigos (ex.: "por", "by", "di").
• Adicionadas opções de indentação (tabulações/espaços com largura) e Tab/Shift+Tab para indentar/desindentar linhas selecionadas.
• Corrigida a limpeza de Markdown: agora trata bullets '*' quando a preservação de listas está desativada.
Correções de bugs
• Corrigido um bug em que audiolivros com SAPI4 podiam ser criados de forma diferente do esperado.
• Janela Buscar em arquivos: ao pressionar Enter em um resultado agora abre na posição correta do trecho e Esc volta aos resultados.
• Janela Opções: ajustado o layout visual das abas Geral, Voz, Editor e Áudio para evitar controles ausentes ou cortados.
• Corrigido um problema de marcadores ao alterar a velocidade de reprodução.
• Corrigido um problema com o Podcast Index e categorias que não eram exibidas corretamente.
• Corrigido o problema do apóstrofo que quebrava a leitura: não há mais leitura separada para diálogos, usam-se tags de voz.

Versão 0.6.3 – 2026-01-30
Melhorias
• Melhorada a detecção do microfone.
• Adicionada reprodução instantânea para todos os formatos.
Correções
• Corrigido o travamento na janela de categorias de podcasts.

Versão 0.6.2 – 2026-01-30
Novas funcionalidades
• Adicionado suporte à execução de arquivos (Shift+F5). Os usuários podem selecionar um interpretador (ex. python) nas Opções, procurá-lo no computador, e pressionando Shift+F5 o script atual é executado. Arquivos HTML abrem no navegador.
• Adicionado suporte para arquivos de ponteiro do Google Docs (.gdoc, .gsheet, .gslides), que abrem automaticamente no navegador padrão.
• Adicionado suporte para o formato de audiolivro M4B (Apple/AAC).
• Adicionada a opção "Mostrar episódios" no menu de contexto dos resultados de pesquisa de podcasts para navegar e reproduzir episódios sem se inscrever.
• Adicionada a função "Ir para linha" (menu Editar ou Ctrl+J) para pular rapidamente para um número de linha específico.
• Adicionadas opções no menu de contexto para ordenar feeds RSS e podcasts (alfabeticamente ou por data).
• Adicionados feeds RSS padrão em vietnamita.
• Adicionada uma caixa de teste de microfone no diálogo de gravação para verificar os níveis antes de começar.
• Adicionada "Mostrar descrição" para episódios de podcast no menu de contexto.
• Adicionado suporte para formatos de áudio/vídeo estendidos via FFmpeg: mkv, avi, mov, m4v, webm, mpg, ts, wmv, flv, vob, 3gp, flac, ogg, wma, aiff.
• Adicionada leitura sincronizada de legendas (srt, vtt, ass, sub, sbv, lrc, smi) com NVDA ou voz selecionada. O programa procura um arquivo de legendas com o mesmo nome do arquivo de mídia. Adicionadas as opções "Importar legendas" e "Remover legendas" no menu Reprodução para arquivos com nomes diferentes.
• Adicionadas associações de arquivos para todos os novos formatos de áudio/vídeo suportados no menu de contexto "Abrir com Sonarpad".
• Adicionada configuração para ajustar o pitch de qualquer arquivo.
• Adicionada opção nas Configurações Gerais para ativar ou desativar relatórios de erros anônimos. Adicionada uma entrada no menu Ajuda para criar um arquivo ZIP de diagnóstico.
• Adicionada opção para usar uma voz diferente para diálogos, tanto para leitura ao vivo quanto para criação de audiolivros.
• Adicionado o navegador de categorias de podcasts para explorar podcasts por categoria (negócios, arte, esportes, etc.).
Melhorias
• Abrir um arquivo de áudio/vídeo do Explorador agora abre diretamente a visualização do reprodutor em vez do editor de texto.
• Removida a solicitação de OCR para PDFs não acessíveis; o OCR agora é realizado automaticamente para melhorar a velocidade e experiência do usuário.
• Melhorado o Terminal Acessível: a leitura NVDA agora lembra a última linha lida para melhor continuidade.
• SAPI 4: A criação de audiolivros agora está totalmente paralelizada e é quase instantânea. Adicionada uma solicitação para escolher o número de processos simultâneos.
• SAPI 4: Eliminado o gargalo WAV-MP3 convertendo fragmentos em paralelo durante a síntese.
• SAPI 4: Melhorado o tratamento de erros e limpeza automática de arquivos temporários.
• Diálogo Localizar: Renomeado "Regex" para "Expressão regular" para maior clareza e adicionadas as traduções ausentes para as opções de pesquisa.
• Audiolivros M4B: Melhor tratamento de saída; dividir por partes/marcadores agora produz um único arquivo M4B com metadados de capítulos incluindo título e autor.
• Reprodutor: Corrigida a precisão de marcadores e anúncios de tempo quando a velocidade de reprodução não é 1.0x.
• Restaurada a navegação Ctrl+Tab e Ctrl+Shift+Tab nas Opções.
• Adicionada uma opção no menu Reprodução para redefinir instantaneamente a velocidade para Normal (1.0x).
• Atualizadas todas as dependências para as versões mais recentes para melhor desempenho e estabilidade.
• Integrado FFmpeg com carregamento dinâmico de DLL para garantir compatibilidade sem bloquear a inicialização.
• Atualizados os filtros de download de podcasts para incluir os novos formatos de áudio/vídeo.
• Impedido que Ctrl+S salve arquivos de áudio/vídeo para evitar corrupção.
• Melhorada a importação de transcrições do YouTube tornando-a mais robusta e resiliente.
• Melhorada a robustez da divisão em partes de audiolivros, garantindo que nenhum texto seja perdido.
• O instalador agora é totalmente multilíngue, suportando Italiano, Inglês, Espanhol, Português, Sueco e Vietnamita com base no idioma do sistema do usuário. O inglês é o padrão para sistemas não suportados.
• Categorias de podcasts: pressionar Enter em uma categoria agora confirma a seleção (equivalente ao botão OK).
• Melhorado o sistema de detecção de travamentos para evitar falsos positivos quando há diálogos modais abertos (mensagens de erro, "texto não encontrado").
Correções
• Corrigido um erro em que o changelog não abria na inicialização.
• Corrigido um erro em que a solicitação de OCR não aparecia para PDFs não acessíveis abertos do Explorador.
• Corrigido um erro de inicialização que podia causar perda de foco ou fechamento de janelas imediatamente após abrir.
• Corrigido um erro crítico na pesquisa regex que impedia encontrar texto, incluindo problemas com "Pesquisa circular" e a opção "Ponto equivale a nova linha" com terminações de linha do Windows.
Localização
• Adicionada a tradução para polonês.
• Adicionada a tradução para francês.
• Adicionada a tradução para tcheco (graças a Radek Žalud e Jiri Holzinger).

Versão 0.6.1 – 2026-01-20
Correções
• Corrigido um erro em que, ao ativar “Mostrar vozes no editor” e reproduzir um podcast, a reprodução era interrompida.
• Corrigido um problema em que alguns podcasts não podiam ser adicionados por URL porque o endereço era truncado.
• Corrigido um erro que impedia a adição de URLs normais na funcionalidade de feeds RSS.
• Corrigido um problema em que o idioma da Wikipedia era mostrado em várias abas das opções.
• Removida a criação de alguns ficheiros de depuração que eram gerados mesmo em modo release.
Melhorias
• Melhorado o suporte para vozes Microsoft, que agora são reproduzidas utilizando um método dedicado com um user agent diferente.
• Adicionado suporte para ficheiros MP4.


Versão 0.6.0 – 2025-01-XX
Melhorias
• Melhorado o suporte para vozes Microsoft, que agora são reproduzidas utilizando um método dedicado com um user agent diferente.


Versão 0.6.0 – 2025-01-20
Novas funcionalidades
• Adicionado o corretor ortográfico. A partir do menu contextual, é possível verificar se a palavra atual está correta e, caso não esteja, obter sugestões.
• Adicionada a importação e exportação de podcasts por meio de arquivos OPML.
• Adicionado suporte à pesquisa no Podcast Index além do iTunes. O utilizador pode introduzir a sua API key e API secret gratuitos (gerados apenas com o seu endereço de e-mail).
• Adicionado suporte às vozes SAPI4, tanto para leitura em tempo real como para a criação de audiolivros
• Adicionado um fallback automático de OCR para PDFs não acessíveis: quando não é encontrado texto extraível, o documento é reconhecido através de OCR..
• Adicionado suporte de dicionário através do Wiktionary. Ao pressionar a tecla Aplicações, são apresentadas as definições e, quando disponíveis, também sinónimos e traduções para outros idiomas.
• Adicionada a importação de artigos da Wikipedia com pesquisa, seleção de resultados e importação direta para o editor.
• Adicionado o atalho Shift+Enter no módulo RSS para abrir um artigo diretamente no site original.
Melhorias
• A seleção do microfone agora é sempre respeitada pela aplicação.
• Na janela de podcasts, ao pressionar Enter num episódio, o NVDA anuncia imediatamente “a carregar”, fornecendo confirmação imediata da ação.
• Nos resultados de pesquisa de podcasts, ao pressionar Enter, o utilizador passa a subscrever o podcast selecionado.
• Corrigidas e melhoradas as etiquetas dos atalhos Ctrl+Shift+O e Podcast Ctrl+Shift+P.
• A velocidade de reprodução e o volume passam agora a ser guardados nas definições e mantêm-se para todos os ficheiros de áudio.
• Adicionada uma pasta de cache dedicada para os episódios de podcasts. O utilizador pode conservar os episódios através de “Conservar podcast” no menu Reproduzir. A cache é limpa automaticamente quando ultrapassa o tamanho definido pelo utilizador (Opções → Áudio).
• Melhorada de forma significativa a obtenção de artigos RSS utilizando libcurl com impersonação de Chrome e iPhone, garantindo compatibilidade com cerca de 99 % dos sites.
• Adicionado o estado lido / não lido para os artigos RSS, com indicação clara na lista RSS.
• A função Substituir tudo agora mostra também o número de substituições efetuadas.
• Adicionado o botão Eliminar podcast ao navegar pela biblioteca de podcasts através da tecla Tab.
Correções
• Removida a entrada redundante “pending update” do menu Ajuda (as atualizações já são geridas automaticamente).
• Corrigido um erro em que, ao abrir um ficheiro MP3 e pressionar Ctrl+S, o ficheiro era guardado e ficava corrompido.
• Corrigido um problema de interface em que “Batch Audiobooks” era apresentado como “(B)… Ctrl+Shift+B” (removida a etiqueta redundante).
• Corrigido o funcionamento das aspas inteligentes: quando ativadas, as aspas normais passam agora a ser corretamente substituídas por aspas tipográficas.
• Corrigido um erro em que, ao utilizar “Ir para o marcador”, a velocidade de reprodução era reposta para 1.0.
• Corrigido um problema em que episódios de podcasts já descarregados eram novamente descarregados em vez de ser utilizada a versão em cache.
Atalhos de teclado
• F1 agora abre o guia.
• F2 agora verifica a existência de atualizações.
• F7 / F8 agora permitem navegar para o erro ortográfico anterior ou seguinte.
• F9 / F10 agora permitem alternar rapidamente entre as vozes guardadas nos favoritos.
Melhorias para desenvolvedores
• Os erros deixaram de ser ignorados silenciosamente: todos os padrões let _ = foram removidos e os erros são agora tratados explicitamente (propagados, registados ou tratados com mecanismos de fallback adequados).
• O projeto agora não compila se existirem avisos: tanto cargo check como cargo clippy devem passar sem warnings, com lints mais restritivos e remoção de allow sempre que possível.
• Removidas implementações personalizadas do tipo strlen / wcslen. Os comprimentos de strings e buffers UTF-16 passam agora a ser derivados de dados geridos pelo Rust, sem varrimentos manuais de memória.
• A gestão de DLL foi limpa e consolidada em torno de libloading, evitando lógica de carregamento personalizada e parsing PE.
• Removidos os helpers manuais de parsing de bytes: todo o parsing passa agora a utilizar from_le_bytes / from_be_bytes sobre slices verificadas.
Estas alterações reduzem o uso desnecessário de unsafe, eliminam potenciais comportamentos indefinidos e tornam a base de código mais idiomática, robusta e fácil de manter.

Versao 0.5.9 - 2025-01-13
Novas funcionalidades
• Adicionada a possibilidade de reordenar RSS pelo menu contextual (cima/baixo/posicao), com validacao de posicoes invalidas.
• Adicionado menu contextual para artigos com abrir site original e compartilhar via WhatsApp, Facebook e X.
• Adicionado atalho Esc para voltar de artigos importados para a lista de RSS.
• Adicionada a modalidade podcast: buscar, inscrever e ouvir; reordenar assinaturas; Esc para parar a reproducao e voltar a lista; Enter em um episodio inicia a reproducao.
• Adicionado controle de velocidade de reproducao para podcasts e arquivos MP3.
• Adicionado Ctrl+T para ir a um tempo especifico.
• Adicionado um botao de previa de voz apos o combo de volume.
• Adicionada a funcao regex para Localizar e Substituir, estilo Notepad++.
• Adicionada a importacao de RSS a partir de arquivos OPML e TXT.
• Adicionada nas Opcoes a caixa para habilitar "Abrir com Sonarpad" no Explorador de arquivos, inclusive na versao portable.
Melhorias
• Melhorada a selecao de velocidade, tom e volume das vozes, respeitando os limites maximos do TTS.
• Varias melhorias no RSS para baixar todos os artigos sem mover o foco do NVDA durante atualizacoes.
• Melhorada a reproducao de audio com um menu dedicado, anuncio de tempo com Ctrl+I e volume ate 300%.
• Adicionados atalhos faltantes para algumas funcoes.
• Reorganizado o menu Editar com um submenu para as funcoes de limpeza de texto.
• Reorganizadas as Opcoes em abas, com Ctrl+Tab e Ctrl+Shift+Tab para navegar.
• Resolvidos os problemas de leitura de artigos: o leitor RSS agora mostra os artigos completos como no navegador.
Correcoes
• Corrigido um problema em que a limpeza de Markdown removia numeros no inicio da linha.
• Corrigido AltGr+Z que acionava Undo.
• Corrigido um problema em que ao gravar um audiolivro nao era possivel interromper rapidamente.
Localizacao
• Adicionada a traducao vietnamita (graças a Anh Duc Nguyen).

Versao 0.5.8 - 2026-01-10
Novas funcionalidades
• Adicionado controle de volume para o microfone e o audio do sistema ao gravar podcasts.
• Adicionada uma nova funcao para importar artigos de sites ou feeds RSS, incluindo os feeds mais importantes para cada idioma.
• Adicionada uma funcao para remover todos os marcadores do arquivo atual.
• Adicionada a funcao para remover linhas duplicadas e linhas duplicadas consecutivas.
• Adicionada a funcao para fechar todas as abas ou janelas exceto a atual.
• Adicionada a entrada Doacoes no menu Ajuda para todos os idiomas.
Melhorias
• Melhorado o terminal acessivel para evitar alguns crashes.
• Melhoradas e corrigidas as access key e os atalhos de teclado do programa.
• Corrigido um problema em que, ao fechar a janela de reproducao de audio, a reproducao nao parava.
• Adicionadas janelas de confirmacao para acoes importantes (ex.: remover linhas duplicadas, remover hifens no fim da linha, remover todos os marcadores do arquivo atual). Nenhuma confirmacao e mostrada se a acao nao se aplica.
• Adicionada a possibilidade de excluir feeds/sites RSS da biblioteca selecionando-os e pressionando Delete.
• Adicionado um menu contextual na janela RSS para modificar ou eliminar feeds/sites RSS.
• Removida a opcao para mover as definicoes para a pasta atual; agora o programa faz isso automaticamente (se a pasta do exe se chama "sonarpad portable" ou o exe esta em unidade removivel, salva na pasta do exe em `config`, senao em `%APPDATA%\\Sonarpad`, com fallback para `config` se a pasta preferida nao for gravavel).

Versao 0.5.7 - 2026-01-05
Novas funcionalidades
• Adicionada a opcao para gravar audiolivros em lote (conversao multipla de arquivos e pastas).
• Adicionado suporte para arquivos Markdown (.md).
• Adicionada a escolha da codificacao ao abrir arquivos de texto.
• Adicionada opcao no terminal para anunciar novas linhas com NVDA.
Melhorias
• A gravacao de audiolivros agora e salva em MP3 nativo quando selecionado.
• O usuario pode escolher onde inserir o asterisco * que indica modificacoes nao salvas.
• Melhorado o sistema de atualizacao para ser mais robusto em diferentes cenarios.
• Adicionada no menu Editar a funcao para remover hifens no final da linha (util para textos OCR).

Versao 0.5.6 - 2026-01-04
Correcoes
  Melhorado Procurar em arquivos: ao pressionar Enter abre o arquivo exatamente no trecho selecionado.
Melhorias
  Suporte a PPT/PPTX.
  Para formatos nao textuais, Salvar agora propoe sempre .txt para evitar corromper a formatacao (PDF/DOC/DOCX/EPUB/HTML/PPT/PPTX).
  Gravacao de podcast do microfone e/ou audio do sistema (menu Arquivo, Ctrl+Shift+R).

Versao 0.5.5 - 2026-01-03
Novas funcionalidades
• Adicionado um terminal acessivel otimizado para muita saida e leitores de tela (Ctrl+Shift+P).
• Adicionada a opcao de guardar as definicoes na pasta atual (modo portable).
Correcoes
• Melhorados os trechos de Procurar em arquivos para manter a previsualizacao alinhada com a ocorrencia.

Versao 0.5.4 – 2026-01-03
Melhorias
• Correcao da funcao Normalizar espacos em branco (Ctrl+Shift+Enter).
• Suporte a HTML/HTM (abrir como texto).

Versao 0.5.3 – 2026-01-02
Novos recursos
• Adicionado Buscar em arquivos.
• Adicionadas novas ferramentas de texto: Normalizar espacos em branco, Quebra de linha dura e Remover Markdown.
• Adicionadas Estatisticas de texto (Alt+Y).
• Adicionados novos comandos de lista no menu Editar:
• Ordenar itens (Alt+Shift+O)
• Manter itens unicos (Alt+Shift+K)
• Inverter itens (Alt+Shift+Z)
• Adicionados Comentar / Descomentar linhas (Ctrl+Q / Ctrl+Shift+Q).
Localizacao
• Adicionada a localizacao em espanhol.
• Adicionada a localizacao em portugues.
Melhorias
• Quando um arquivo EPUB esta aberto, Salvar muda automaticamente para Salvar como e exporta o conteudo como .txt para evitar corromper o EPUB.

## 0.5.2 - 2026-01-01
- Adicionado um changelog.
- Adicionadas opcoes "Abrir com Sonarpad" e associacoes de arquivos suportados durante a instalacao.
- Melhorada a localizacao de mensagens (erros, dialogos, exportacao de audiolivro).
- Adicionada a selecao de partes ao usar "Dividir audiolivro por texto", com a opcao "Exigir o marcador no inicio da linha".
- Adicionada a importacao de transcricoes do YouTube com selecao de idioma, opcao de timestamps e melhorias de foco.

## 0.5.1 - 2025-12-31
- Atualizacoes automaticas com confirmacao, melhorias de erros e notificacoes.
- Melhorias na exportacao de audiolivros (divisao por texto, SAPI5/Media Foundation, controles avancados).
- Melhorias em TTS (pausa/retomar, dicionario de substituicoes, favoritos).
- Menu Ver e paineis de vozes/favoritos, cor e tamanho de texto.
- Idioma padrao do sistema e melhorias de localizacao.
- CI e empacotamento Windows (artefatos, MSI/NSIS, cache).

## 0.5.0 - 2025-12-27
- Refatoracao modular (editor, manipulacao de arquivos, menu, busca).
- Workflow de compilacao/empacotamento Windows e atualizacoes de README/licenca.
- Correcao da navegacao TAB na janela de Ajuda.

## 0.5 - 2025-12-27
- Atualizacao preliminar da versao.

## 0.1.0 - 2025-12-25
- Versao inicial: estrutura do projeto e README.











