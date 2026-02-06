# Changelog

Versão 0.6.5 – 05/02/2026
Melhorias
• Tradução em espanhol aprimorada graças a Arturo Fernandez Rivas.
Correções de bugs
• A leitura a partir do cursor (F5) agora começa exatamente no cursor. Antes podia começar algumas linhas acima porque o deslocamento do cursor não correspondia às posições CRLF/UTF-16.
• Corrigido um problema de redesenho: ao digitar sobre uma seleção, o texto anterior podia desaparecer até mover a seleção.
• Corrigida a falha ao dividir por tempo audiolivros a partir de EPUB: o Edge TTS podia falhar com chunks vazios ou muito longos ("Edge audio not sent").

Versão 0.6.4 – 05/02/2026
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
