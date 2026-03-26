library(shiny)
library(bs4Dash)
library(dplyr)
library(ggplot2)
library(DT)
library(stringr)
library(tibble)
library(ggchicklet)
library(grid)
library(scales)
library(readr)
library(tidyr)

carregar_dados <- function() {
  googlesheets4::gs4_deauth()

  tryCatch(
    googlesheets4::read_sheet(
      "https://docs.google.com/spreadsheets/d/1__Nu1bVnEOxgNFOayqDqCa099oPl54v3YptvoT7KzHY/edit?usp=sharing"
    ),
    error = function(e) {
      stop(
        paste("Falha ao carregar a planilha do Google Sheets:", conditionMessage(e)),
        call. = FALSE
      )
    }
  )
}

dados <- carregar_dados() |>
  mutate(
    repeticao = suppressWarnings(as.integer(repeticao)),
    produtividade = suppressWarnings(as.numeric(produtividade)),
    media_prod = suppressWarnings(as.numeric(media_prod))
  )

gerar_custo_ficticio <- function(df) {
  dose_num <- suppressWarnings(readr::parse_number(as.character(df$dose)))
  dose_num[is.na(dose_num)] <- 0

  produto_chr <- stringr::str_to_upper(trimws(as.character(df$produto)))
  aplicacao_chr <- trimws(as.character(df$aplicacao))

  custo_base <- 3.5 +
    pmin(dose_num * 0.9, 18) +
    ifelse(!is.na(aplicacao_chr) & aplicacao_chr != "", 1.2, 0) +
    ifelse(stringr::str_detect(produto_chr, "FUNG"), 3.2, 0) +
    ifelse(stringr::str_detect(produto_chr, "HERB"), 2.6, 0) +
    ifelse(stringr::str_detect(produto_chr, "INSET"), 2.1, 0) +
    ifelse(stringr::str_detect(produto_chr, "BIO"), 1.4, 0)

  custo_base <- round(custo_base, 1)
  custo_base[produto_chr == "TESTEMUNHA" | is.na(produto_chr) | produto_chr == ""] <- 0

  dplyr::mutate(df, custo_sc_ha = custo_base)
}

dados <- gerar_custo_ficticio(dados)

coluna_produtividade <- if ("media_prod" %in% names(dados)) "media_prod" else "produtividade"

colunas_indicadores <- c(
  "ferrugem", "manchas", "oidio",
  "buva", "caruru", "poaia",
  "percevejos", "lagartas", "tripes",
  "vigor", "altura", "stand",
  "ndvi", "ndre", "pmg"
)

colunas_indicadores <- intersect(colunas_indicadores, names(dados))

opcao_vazia <- function(valores) {
  c("", sort(stats::na.omit(unique(valores))))
}

ordenar_tratamentos <- function(df) {
  df |>
    filter(!is.na(tratamento), tratamento != "") |>
    distinct(tratamento) |>
    mutate(ordem = readr::parse_number(tratamento)) |>
    arrange(ordem, tratamento) |>
    pull(tratamento)
}

filtrar_base <- function(
  df,
  safra = NULL,
  cultura = NULL,
  segmento = NULL,
  ensaio = NULL,
  tratamentos = NULL
) {
  if (!is.null(safra) && nzchar(safra)) {
    df <- filter(df, safra == !!safra)
  }

  if (!is.null(cultura) && nzchar(cultura)) {
    df <- filter(df, cultura == !!cultura)
  }

  if (!is.null(segmento) && nzchar(segmento)) {
    df <- filter(df, segmento == !!segmento)
  }

  if (!is.null(ensaio) && nzchar(ensaio)) {
    df <- filter(df, ensaio == !!ensaio)
  }

  if (!is.null(tratamentos)) {
    if (length(tratamentos) == 0) {
      return(df[0, , drop = FALSE])
    }

    df <- filter(df, tratamento %in% tratamentos)
  }

  df
}

indicadores_disponiveis <- function(df) {
  if (nrow(df) == 0) {
    return(character(0))
  }

  colunas_indicadores[
    vapply(
      colunas_indicadores,
      function(coluna) any(!is.na(df[[coluna]])),
      logical(1)
    )
  ]
}

resumo_tratamentos <- function(df, coluna_prod) {
  if (nrow(df) == 0) {
    return(tibble(
      tratamento = character(),
      produto = character(),
      dose = character(),
      aplicacao = character(),
      n = integer(),
      media_prod = numeric(),
      sd_prod = numeric(),
      cv_prod = numeric(),
      scott_knott = character(),
      ganho_vs_testemunha = numeric()
    ))
  }

  media_testemunha <- df |>
    filter(str_to_upper(produto) == "TESTEMUNHA") |>
    summarise(media = mean(.data[[coluna_prod]], na.rm = TRUE)) |>
    pull(media)

  df |>
    group_by(tratamento, produto, dose, aplicacao) |>
    summarise(
      n = n(),
      media_prod = mean(.data[[coluna_prod]], na.rm = TRUE),
      sd_prod = sd(.data[[coluna_prod]], na.rm = TRUE),
      cv_prod = ifelse(media_prod == 0, NA_real_, sd_prod / media_prod * 100),
      scott_knott = {
        grupos <- na.omit(scott_knott)
        if (length(grupos) == 0) NA_character_ else as.character(grupos[1])
      },
      .groups = "drop"
    ) |>
    mutate(
      ganho_vs_testemunha = ifelse(
        is.na(media_testemunha) || media_testemunha == 0,
        NA_real_,
        (media_prod - media_testemunha) / media_testemunha * 100
      )
    ) |>
    arrange(desc(media_prod))
}

formatar_kpi <- function(valor) {
  HTML(sprintf("<span style='font-size:34px; font-weight:700;'>%s</span>", valor))
}

rotulo_filtro <- function(valor, vazio = "Todos") {
  if (is.null(valor) || length(valor) == 0) {
    return(vazio)
  }

  valor <- valor[!is.na(valor) & nzchar(valor)]

  if (length(valor) == 0) {
    return(vazio)
  }

  paste(valor, collapse = ", ")
}

quebrar_em_linhas <- function(texto, largura = 95) {
  if (length(texto) == 0 || is.na(texto) || !nzchar(texto)) {
    return("")
  }

  paste(strwrap(texto, width = largura), collapse = "\n")
}

adicionar_pagina_texto_pdf <- function(
  titulo,
  linhas,
  subtitulo = NULL,
  familia = "mono",
  tamanho = 0.8
) {
  grid::grid.newpage()
  grid::pushViewport(grid::viewport(layout = grid::grid.layout(3, 1, heights = unit(c(1.2, 0.7, 10), "null"))))

  grid::grid.text(
    titulo,
    x = unit(0.02, "npc"),
    y = unit(0.5, "npc"),
    just = c("left", "center"),
    gp = grid::gpar(fontsize = 18, fontface = "bold", col = "#0A1F44"),
    vp = grid::viewport(layout.pos.row = 1, layout.pos.col = 1)
  )

  if (!is.null(subtitulo) && nzchar(subtitulo)) {
    grid::grid.text(
      subtitulo,
      x = unit(0.02, "npc"),
      y = unit(0.9, "npc"),
      just = c("left", "top"),
      gp = grid::gpar(fontsize = 10, col = "#555555"),
      vp = grid::viewport(layout.pos.row = 2, layout.pos.col = 1)
    )
  }

  grid::grid.text(
    paste(linhas, collapse = "\n"),
    x = unit(0.02, "npc"),
    y = unit(0.98, "npc"),
    just = c("left", "top"),
    gp = grid::gpar(fontfamily = familia, fontsize = tamanho * 12, col = "#1A1A1A"),
    vp = grid::viewport(layout.pos.row = 3, layout.pos.col = 1)
  )
}

adicionar_paginas_tabela_pdf <- function(df, titulo, subtitulo = NULL, linhas_por_pagina = 24) {
  if (nrow(df) == 0) {
    adicionar_pagina_texto_pdf(
      titulo = titulo,
      subtitulo = subtitulo,
      linhas = "Sem dados para os filtros selecionados.",
      familia = "sans",
      tamanho = 1
    )
    return(invisible(NULL))
  }

  total_paginas <- ceiling(nrow(df) / linhas_por_pagina)

  for (pagina in seq_len(total_paginas)) {
    inicio <- ((pagina - 1) * linhas_por_pagina) + 1
    fim <- min(pagina * linhas_por_pagina, nrow(df))
    bloco <- df[inicio:fim, , drop = FALSE]

    linhas_tabela <- capture.output(
      print.data.frame(bloco, row.names = FALSE, right = FALSE)
    )

    adicionar_pagina_texto_pdf(
      titulo = titulo,
      subtitulo = paste0(
        subtitulo,
        if (!is.null(subtitulo) && nzchar(subtitulo)) " | " else "",
        "Pagina ", pagina, " de ", total_paginas
      ),
      linhas = linhas_tabela
    )
  }

  invisible(NULL)
}

plot_relatorio_produtividade <- function(df_plot, indicador_nome) {
  if (nrow(df_plot) == 0) {
    return(NULL)
  }

  ggplot(df_plot, aes(x = tratamento, y = produtividade)) +
    geom_col(fill = "#0A1F44", width = 0.7) +
    geom_errorbar(
      aes(ymin = produtividade - se * 2, ymax = produtividade + se * 2),
      width = 0.2,
      linewidth = 0.35,
      colour = "#444444"
    ) +
    geom_line(aes(y = avaliacao, group = 1), linewidth = 1.2, colour = "#D4A017") +
    geom_point(aes(y = avaliacao), size = 2.8, colour = "#D4A017") +
    geom_text(
      aes(y = produtividade, label = round(produtividade, 1)),
      vjust = -0.4,
      size = 3.4,
      fontface = "bold",
      colour = "#0A1F44"
    ) +
    labs(
      title = "Produtividade por tratamento",
      subtitle = paste("Indicador:", indicador_nome),
      x = "Tratamento",
      y = "Produtividade media (sc/ha)",
      caption = "Linha dourada: media da avaliacao selecionada."
    ) +
    theme_minimal(base_size = 12) +
    theme(
      panel.grid.major.x = element_blank(),
      panel.grid.minor = element_blank(),
      axis.text.x = element_text(face = "bold"),
      plot.title = element_text(face = "bold", colour = "#0A1F44")
    )
}

plot_relatorio_correlacao <- function(df_plot, indicador_nome) {
  if (nrow(df_plot) <= 1) {
    return(NULL)
  }

  ggplot(df_plot, aes(x = avaliacao_plot, y = produtividade_plot)) +
    geom_point(size = 2.8, alpha = 0.85, colour = "#0A1F44") +
    geom_smooth(method = "lm", se = TRUE, colour = "#D4A017", fill = "#F2D77A", linewidth = 1) +
    labs(
      title = "Correlacao entre produtividade e avaliacao",
      x = indicador_nome,
      y = "Produtividade (sc/ha)"
    ) +
    theme_minimal(base_size = 12) +
    theme(
      plot.title = element_text(face = "bold", colour = "#0A1F44")
    )
}

ui <- bs4Dash::dashboardPage(
  header = bs4Dash::dashboardHeader(
    title = tags$div(
      tags$img(src = "agrometric2.jpg", height = "85px", style = "margin-right:5px;"),
      span("Painel de Ensaios")
    )
  ),
  sidebar = bs4Dash::dashboardSidebar(
    skin = "light",
    status = "navy",
    title = "Filtros",
    style = "height: calc(100vh - 60px); overflow-y: auto;",
    bs4Dash::sidebarMenu(
      id = "tabs",
      bs4Dash::menuItem("Visao Geral", tabName = "visao_geral", icon = shiny::icon("chart-line")),
      bs4Dash::menuItem("Resultados", tabName = "resultados_tab", icon = shiny::icon("seedling")),
      bs4Dash::menuItem("Indicadores", tabName = "indicadores_tab", icon = shiny::icon("microscope")),
      bs4Dash::menuItem("Tabela Resumo", tabName = "tabela_tab", icon = shiny::icon("table"))
    ),
    tags$hr(),
    selectizeInput("filtro_safra", "Safra", choices = c("", "carregando"), selected = ""),
    selectizeInput("filtro_cultura", "Cultura", choices = c("", "carregando"), selected = ""),
    selectizeInput("filtro_segmento", "Segmento", choices = c("", "carregando"), selected = ""),
    conditionalPanel(
      condition = "input.tabs != 'visao_geral'",
      uiOutput("ui_filtro_ensaio")
    ),
    conditionalPanel(
      condition = "input.tabs == 'resultados_tab' || input.tabs == 'indicadores_tab'",
      selectizeInput(
        "filtro_tratamento",
        "Tratamento",
        choices = NULL,
        selected = NULL,
        multiple = TRUE
      )
    ),
    conditionalPanel(
      condition = "input.tabs == 'resultados_tab' || input.tabs == 'indicadores_tab'",
      tags$hr(),
      uiOutput("ui_indicador"),
      checkboxInput("ordenar_indicador", "Ordenar pela avaliacao", value = FALSE)
    ),
    tags$br(),
    div(
      style = "text-align:center; margin-top: 8px;",
      actionButton(
        "limpar_filtros",
        "Limpar filtros",
        icon = shiny::icon("eraser"),
        width = "85%",
        class = "btn-warning"
      )
    ),
    tags$br(),
    div(
      style = "text-align:center; margin-top: 4px;",
      downloadButton(
        "baixar_relatorio_pdf",
        "Baixar relatorio PDF",
        class = "btn-primary",
        width = "85%"
      )
    ),
    tags$br(),
    tags$br(),
  ),
  body = bs4Dash::dashboardBody(
    tags$head(
      tags$link(rel = "icon", type = "image/jpg", href = "agrometric2.jpg"),
      tags$style(HTML("
        .small-box .inner h3 {
          font-size: 18px !important;
          font-weight: bold !important;
        }
        .small-box .inner p {
          font-size: 18px !important;
          font-weight: 400 !important;
        }
        .small-box .icon {
          color: rgba(100, 100, 100, 0.42) !important;
        }
        .selectize-control.single .selectize-input,
        .selectize-control.multi .selectize-input {
          min-height: 34px !important;
          padding: 4px 8px !important;
          font-size: 15px !important;
        }
        .form-group {
          margin-bottom: 1px !important;
        }
        .control-label {
          font-size: 15px !important;
          margin-bottom: 1px !important;
        }
      "))
    ),
    bs4Dash::tabItems(
      bs4Dash::tabItem(
        tabName = "visao_geral",
        fluidRow(
          bs4Dash::valueBoxOutput("kpi_ensaios", width = 3),
          bs4Dash::valueBoxOutput("kpi_tratamentos", width = 3),
          bs4Dash::valueBoxOutput("kpi_total_parcelas", width = 3),
          bs4Dash::valueBoxOutput("kpi_media_repeticao", width = 3)
        ),
        fluidRow(
          bs4Dash::box(
            title = "Visao geral da area experimental",
            width = 12,
            solidHeader = TRUE,
            plotOutput("graf_ranking", height = 420)
          )
        ),
        fluidRow(
          bs4Dash::valueBoxOutput("kpi_media_prod", width = 3),
          bs4Dash::valueBoxOutput("kpi_range_prod", width = 3),
          bs4Dash::valueBoxOutput("kpi_ganho_test", width = 3),
          bs4Dash::valueBoxOutput("kpi_media_cv", width = 3)
        )
      ),
      bs4Dash::tabItem(
        tabName = "resultados_tab",
        fluidRow(
          bs4Dash::box(
            width = 12,
            title = "Comparativo de Produtividade vs Avaliacao",
            status = "warning",
            solidHeader = TRUE,
            collapsible = TRUE,
            plotOutput("graf_trat_indicador", height = "520px")
          )
        ),
        fluidRow(
          bs4Dash::box(
            width = 12,
            title = "Descricao dos Tratamentos do Ensaio",
            status = "navy",
            solidHeader = TRUE,
            collapsible = TRUE,
            DT::DTOutput("tabela_descricao_ensaio")
          )
        )
      ),
      bs4Dash::tabItem(
        tabName = "indicadores_tab",
        fluidRow(
          bs4Dash::valueBoxOutput("kpi_cor_indicador", width = 4),
          bs4Dash::valueBoxOutput("kpi_r2_indicador", width = 4),
          bs4Dash::valueBoxOutput("kpi_obs_indicador", width = 4)
        ),
        fluidRow(
          bs4Dash::box(
            title = "Media da avaliacao por tratamento",
            width = 6,
            solidHeader = TRUE,
            plotOutput("graf_indicador", height = 420)
          ),
          bs4Dash::box(
            title = "Correlacao entre produtividade e avaliacao",
            width = 6,
            solidHeader = TRUE,
            plotOutput("graf_correlacao_indicador", height = 420)
          )
        ),
        fluidRow(
          bs4Dash::box(
            title = "Distribuicao de produtividade e avaliacao",
            width = 12,
            solidHeader = TRUE,
            plotOutput("graf_boxplots_indicadores", height = 480)
          )
        )
      ),
      bs4Dash::tabItem(
        tabName = "tabela_tab",
        fluidRow(
          bs4Dash::box(
            title = "Resumo tecnico por tratamento",
            width = 12,
            solidHeader = TRUE,
            DT::DTOutput("tabela_resumo")
          )
        )
      )
    )
  ),
  controlbar = bs4Dash::dashboardControlbar(
    title = "Sobre",
    p("Painel inicial em R Shiny + bs4Dash para ensaios agricolas.")
  ),
  footer = bs4Dash::dashboardFooter(
    left = "AgroMetric",
    right = "Painel experimental"
  ),
  title = "Painel de Ensaios"
)

server <- function(input, output, session) {
  observeEvent(TRUE, {
    updateSelectizeInput(session, "filtro_safra", choices = opcao_vazia(dados$safra), selected = "", server = TRUE)
    updateSelectizeInput(session, "filtro_cultura", choices = opcao_vazia(dados$cultura), selected = "", server = TRUE)
    updateSelectizeInput(session, "filtro_segmento", choices = opcao_vazia(dados$segmento), selected = "", server = TRUE)
  }, once = TRUE)

  base_filtrada <- reactive({
    filtrar_base(
      df = dados,
      safra = input$filtro_safra,
      cultura = input$filtro_cultura,
      segmento = input$filtro_segmento,
      ensaio = input$filtro_ensaio
    )
  })

  base_filtrada_tratamentos <- reactive({
    filtrar_base(
      df = base_filtrada(),
      tratamentos = input$filtro_tratamento
    )
  })

  resumo_base <- reactive({
    resumo_tratamentos(base_filtrada(), coluna_produtividade)
  })

  output$ui_filtro_ensaio <- renderUI({
    ensaios_validos <- opcao_vazia(base_filtrada()$ensaio)
    ensaio_atual <- if (!is.null(input$filtro_ensaio) && input$filtro_ensaio %in% ensaios_validos) input$filtro_ensaio else ""

    selectizeInput(
      "filtro_ensaio",
      "Ensaio",
      choices = ensaios_validos,
      selected = ensaio_atual,
      multiple = FALSE
    )
  })

  observeEvent(input$limpar_filtros, {
    updateSelectizeInput(session, "filtro_safra", selected = "")
    updateSelectizeInput(session, "filtro_cultura", selected = "")
    updateSelectizeInput(session, "filtro_segmento", selected = "")
    updateSelectizeInput(session, "filtro_ensaio", selected = "")
    updateSelectizeInput(session, "filtro_tratamento", selected = character(0))
  })

  observeEvent(
    list(input$filtro_safra, input$filtro_cultura, input$filtro_segmento, input$filtro_ensaio, input$tabs),
    {
      req(input$tabs %in% c("resultados_tab", "indicadores_tab"))

      tratamentos_disp <- ordenar_tratamentos(base_filtrada())
      atuais <- isolate(input$filtro_tratamento)

      if (is.null(atuais)) {
        atuais <- character(0)
      }

      selecionados <- intersect(atuais, tratamentos_disp)

      if (length(selecionados) == 0) {
        selecionados <- tratamentos_disp
      }

      updateSelectizeInput(
        session,
        "filtro_tratamento",
        choices = tratamentos_disp,
        selected = selecionados,
        server = TRUE
      )
    },
    ignoreInit = FALSE
  )

  output$ui_indicador <- renderUI({
    cols_ok <- indicadores_disponiveis(base_filtrada())

    selectInput(
      "indicador_escolhido",
      "Avaliacao",
      choices = cols_ok,
      selected = if (length(cols_ok) > 0) cols_ok[1] else NULL
    )
  })

  dados_resultado <- reactive({
    validate(
      need(!is.null(input$filtro_ensaio) && nzchar(input$filtro_ensaio), "Selecione um ensaio."),
      need(!is.null(input$indicador_escolhido) && nzchar(input$indicador_escolhido), "Selecione uma avaliacao."),
      need(!is.null(input$filtro_tratamento) && length(input$filtro_tratamento) > 0, "Selecione ao menos um tratamento.")
    )

    df <- base_filtrada_tratamentos() |>
      mutate(
        tratamento = trimws(as.character(tratamento)),
        produto = trimws(as.character(produto)),
        dose = trimws(as.character(dose)),
        aplicacao = trimws(as.character(aplicacao))
      )

    validate(
      need(nrow(df) > 0, "Os filtros selecionados nao retornaram dados.")
    )

    df |>
      group_by(tratamento) |>
      summarise(
        produtividade = mean(.data[[coluna_produtividade]], na.rm = TRUE),
        sd = sd(.data[[coluna_produtividade]], na.rm = TRUE),
        n = n(),
        se = if_else(n > 1, sd / sqrt(n), 0),
        grupo = first(scott_knott),
        avaliacao = mean(.data[[input$indicador_escolhido]], na.rm = TRUE),
        .groups = "drop"
      ) |>
      filter(is.finite(produtividade), is.finite(avaliacao)) |>
      arrange(
        if (isTRUE(input$ordenar_indicador)) desc(avaliacao) else desc(produtividade)
      ) |>
      mutate(tratamento = factor(tratamento, levels = tratamento))
  })

  dados_indicador <- reactive({
    validate(
      need(!is.null(input$indicador_escolhido) && nzchar(input$indicador_escolhido), "Selecione uma avaliacao."),
      need(!is.null(input$filtro_tratamento) && length(input$filtro_tratamento) > 0, "Selecione ao menos um tratamento.")
    )

    base_filtrada_tratamentos() |>
      mutate(tratamento = trimws(as.character(tratamento))) |>
      group_by(tratamento) |>
      summarise(
        avaliacao = mean(.data[[input$indicador_escolhido]], na.rm = TRUE),
        .groups = "drop"
      ) |>
      filter(is.finite(avaliacao)) |>
      arrange(
        if (isTRUE(input$ordenar_indicador)) desc(avaliacao) else readr::parse_number(tratamento),
        tratamento
      ) |>
      mutate(tratamento = factor(tratamento, levels = tratamento))
  })

  dados_indicador_detalhe <- reactive({
    validate(
      need(!is.null(input$indicador_escolhido) && nzchar(input$indicador_escolhido), "Selecione uma avaliacao."),
      need(!is.null(input$filtro_tratamento) && length(input$filtro_tratamento) > 0, "Selecione ao menos um tratamento.")
    )

    df <- base_filtrada_tratamentos() |>
      mutate(
        tratamento = trimws(as.character(tratamento)),
        produtividade_plot = .data[[coluna_produtividade]],
        avaliacao_plot = .data[[input$indicador_escolhido]]
      ) |>
      filter(is.finite(produtividade_plot), is.finite(avaliacao_plot))

    validate(
      need(nrow(df) > 1, "Dados insuficientes para analise de correlacao e distribuicao.")
    )

    df
  })

  resumo_correlacao_indicador <- reactive({
    df <- dados_indicador_detalhe()

    cor_val <- suppressWarnings(stats::cor(df$produtividade_plot, df$avaliacao_plot, use = "complete.obs"))

    if (!is.finite(cor_val)) {
      cor_val <- NA_real_
    }

    tibble(
      correlacao = cor_val,
      r2 = ifelse(is.na(cor_val), NA_real_, cor_val^2),
      observacoes = nrow(df)
    )
  })

  descricao_ensaio <- reactive({
    validate(
      need(!is.null(input$filtro_ensaio) && nzchar(input$filtro_ensaio), "Selecione um ensaio."),
      need(!is.null(input$filtro_tratamento) && length(input$filtro_tratamento) > 0, "Selecione ao menos um tratamento.")
    )

    base_filtrada_tratamentos() |>
      mutate(
        tratamento = trimws(as.character(tratamento)),
        produto = trimws(as.character(produto)),
        dose = trimws(as.character(dose)),
        aplicacao = trimws(as.character(aplicacao))
      ) |>
      select(tratamento, produto, dose, aplicacao) |>
      distinct() |>
      mutate(ordem = readr::parse_number(tratamento)) |>
      arrange(ordem, tratamento) |>
      select(-ordem)
  })

  output$kpi_ensaios <- bs4Dash::renderValueBox({
    base <- base_filtrada()
    valor <- if (nrow(base) == 0) 0 else base |>
      filter(!is.na(safra), !is.na(cultura), !is.na(ensaio)) |>
      distinct(safra, cultura, ensaio) |>
      nrow()

    bs4Dash::valueBox(formatar_kpi(valor), "Ensaios", icon = icon("flask"), color = "warning")
  })

  output$kpi_tratamentos <- bs4Dash::renderValueBox({
    base <- base_filtrada()
    valor <- if (nrow(base) == 0) 0 else base |>
      filter(!is.na(safra), !is.na(cultura), !is.na(ensaio), !is.na(tratamento)) |>
      distinct(safra, cultura, ensaio, tratamento) |>
      nrow()

    bs4Dash::valueBox(formatar_kpi(valor), "Tratamentos", icon = icon("vials"), color = "warning")
  })

  output$kpi_total_parcelas <- bs4Dash::renderValueBox({
    base <- base_filtrada()
    valor <- if (nrow(base) == 0) 0 else base |>
      filter(!is.na(safra), !is.na(cultura), !is.na(ensaio), !is.na(tratamento), !is.na(repeticao)) |>
      distinct(safra, cultura, ensaio, tratamento, repeticao) |>
      nrow()

    bs4Dash::valueBox(formatar_kpi(valor), "Total de parcelas", icon = icon("th"), color = "warning")
  })

  output$kpi_media_repeticao <- bs4Dash::renderValueBox({
    base <- base_filtrada()
    valor <- if (nrow(base) == 0) {
      NA_real_
    } else {
      base |>
        filter(!is.na(safra), !is.na(cultura), !is.na(ensaio), !is.na(tratamento), !is.na(repeticao)) |>
        group_by(safra, cultura, ensaio, tratamento) |>
        summarise(n_rep = n_distinct(repeticao), .groups = "drop") |>
        summarise(media = mean(n_rep, na.rm = TRUE)) |>
        pull(media)
    }

    bs4Dash::valueBox(
      formatar_kpi(ifelse(is.na(valor), "-", round(valor, 1))),
      "Numero de repeticoes",
      icon = icon("redo"),
      color = "warning"
    )
  })

  output$kpi_media_prod <- bs4Dash::renderValueBox({
    base <- base_filtrada()
    valor <- if (nrow(base) == 0) NA_real_ else mean(base[[coluna_produtividade]], na.rm = TRUE)

    bs4Dash::valueBox(
      formatar_kpi(ifelse(is.na(valor), "-", paste0(round(valor, 1), " sc/ha"))),
      "Produtividade media",
      icon = icon("seedling"),
      color = "navy"
    )
  })

  output$kpi_range_prod <- bs4Dash::renderValueBox({
    base <- base_filtrada()

    valor <- if (nrow(base) == 0) {
      NA_real_
    } else {
      prod_validas <- base[[coluna_produtividade]]
      prod_validas <- prod_validas[!is.na(prod_validas)]

      if (length(prod_validas) == 0) NA_real_ else max(prod_validas) - min(prod_validas)
    }

    bs4Dash::valueBox(
      formatar_kpi(ifelse(is.na(valor), "-", paste0(round(valor, 1), " sc/ha"))),
      "Range produtividade",
      icon = icon("arrows-alt-h"),
      color = "navy"
    )
  })

  output$kpi_ganho_test <- bs4Dash::renderValueBox({
    ganho <- if (nrow(resumo_base()) == 0) {
      NA_real_
    } else {
      resumo_base() |>
        filter(str_to_upper(produto) != "TESTEMUNHA") |>
        summarise(media = mean(ganho_vs_testemunha, na.rm = TRUE)) |>
        pull(media)
    }

    if (length(ganho) == 0 || is.nan(ganho)) {
      ganho <- NA_real_
    }

    bs4Dash::valueBox(
      formatar_kpi(ifelse(is.na(ganho), "-", paste0(round(ganho, 1), "%"))),
      "Ganho medio vs testemunha",
      icon = icon("rocket"),
      color = "navy"
    )
  })

  output$kpi_media_cv <- bs4Dash::renderValueBox({
    base <- base_filtrada()
    valor <- if (nrow(base) == 0 || !("cv_perc" %in% names(base))) NA_real_ else mean(base$cv_perc, na.rm = TRUE)

    if (is.nan(valor)) {
      valor <- NA_real_
    }

    bs4Dash::valueBox(
      formatar_kpi(ifelse(is.na(valor), "-", paste0(round(valor, 1), "%"))),
      "CV medio",
      icon = icon("percentage"),
      color = "navy"
    )
  })

  output$kpi_cor_indicador <- bs4Dash::renderValueBox({
    resumo_cor <- resumo_correlacao_indicador()
    valor <- resumo_cor$correlacao[[1]]

    bs4Dash::valueBox(
      formatar_kpi(ifelse(is.na(valor), "-", round(valor, 2))),
      "Correlacao (r)",
      icon = icon("project-diagram"),
      color = "info"
    )
  })

  output$kpi_r2_indicador <- bs4Dash::renderValueBox({
    resumo_cor <- resumo_correlacao_indicador()
    valor <- resumo_cor$r2[[1]]

    bs4Dash::valueBox(
      formatar_kpi(ifelse(is.na(valor), "-", round(valor, 2))),
      "R2 da relacao",
      icon = icon("chart-area"),
      color = "info"
    )
  })

  output$kpi_obs_indicador <- bs4Dash::renderValueBox({
    resumo_cor <- resumo_correlacao_indicador()
    valor <- resumo_cor$observacoes[[1]]

    bs4Dash::valueBox(
      formatar_kpi(ifelse(is.na(valor), "-", valor)),
      "Observacoes validas",
      icon = icon("database"),
      color = "info"
    )
  })

  output$graf_ranking <- renderPlot({
    base <- base_filtrada()

    if (nrow(base) == 0) {
      return(NULL)
    }

    df_plot <- base |>
      filter(
        !is.na(segmento),
        !is.na(.data[[coluna_produtividade]]),
        !is.na(custo_sc_ha)
      ) |>
      group_by(segmento) |>
      summarise(
        media_prod = mean(.data[[coluna_produtividade]], na.rm = TRUE),
        custo = mean(custo_sc_ha, na.rm = TRUE),
        .groups = "drop"
      ) |>
      arrange(media_prod) |>
      mutate(produtividade_total = media_prod)

    if (nrow(df_plot) == 0) {
      return(NULL)
    }

    ordem_segmento <- df_plot$segmento

    df_azul <- df_plot |>
      transmute(
        segmento = factor(segmento, levels = ordem_segmento),
        valor = produtividade_total,
        legenda = factor(
          "Prod. media (sc/ha)",
          levels = c("Prod. media (sc/ha)", "Custo (sc/ha)")
        )
      )

    df_dourado <- df_plot |>
      transmute(
        segmento = factor(segmento, levels = ordem_segmento),
        valor = custo,
        legenda = factor(
          "Custo (sc/ha)",
          levels = c("Prod. media (sc/ha)", "Custo (sc/ha)")
        )
      )

    ggplot() +
      ggchicklet::geom_chicklet(
        data = df_plot |> mutate(segmento = factor(segmento, levels = ordem_segmento)),
        aes(x = segmento, y = media_prod),
        fill = "grey70",
        width = 0.68,
        radius = grid::unit(5, "pt"),
        alpha = 0.18,
        position = ggplot2::position_nudge(x = 0.03, y = 0.12)
      ) +
      ggchicklet::geom_chicklet(
        data = df_azul,
        aes(x = segmento, y = valor, fill = legenda),
        width = 0.65,
        radius = grid::unit(5, "pt"),
        color = "#08172F"
      ) +
      ggchicklet::geom_chicklet(
        data = df_dourado,
        aes(x = segmento, y = valor, fill = legenda),
        width = 0.65,
        radius = grid::unit(5, "pt"),
        position = ggplot2::position_stack(vjust = 1)
      ) +
      ggplot2::geom_text(
        data = df_plot |> mutate(segmento = factor(segmento, levels = ordem_segmento)),
        aes(x = segmento, y = media_prod * 1.05, label = round(media_prod, 1)),
        vjust = 1,
        color = "#0A1F44",
        fontface = "bold",
        size = 4.5
      ) +
      ggplot2::geom_text(
        data = df_dourado,
        aes(x = segmento, y = valor * 1.35, label = round(valor, 1)),
        vjust = 1,
        color = "white",
        fontface = "bold",
        size = 4.2
      ) +
      scale_fill_manual(
        breaks = c("Prod. media (sc/ha)", "Custo (sc/ha)"),
        values = c("#0A1F44", "#D4AF37")
      ) +
      labs(
        x = NULL,
        y = "Produtividade (sc/ha)",
        title = "Produtividade e custo medio por segmento",
        fill = NULL
      ) +
      theme_minimal(base_size = 15) +
      theme(
        panel.grid.major.x = element_blank(),
        panel.grid.minor = element_blank(),
        panel.grid.major.y = element_line(color = "grey90"),
        axis.title.x = element_blank(),
        axis.text = element_text(color = "#0A1F44"),
        plot.title = element_text(face = "bold", size = 18, color = "#0A1F44", hjust = 0.5),
        legend.position = "right",
        legend.title = element_blank()
      )
  })

  output$graf_indicador <- renderPlot({
    req(input$tabs == "indicadores_tab")
    df_plot <- dados_indicador()

    validate(
      need(nrow(df_plot) > 0, "Sem dados para exibir o indicador.")
    )

    ggplot(df_plot, aes(x = tratamento, y = avaliacao)) +
      ggchicklet::geom_chicklet(
        fill = "#163A5F",
        color = "#08172F",
        width = 0.65,
        radius = grid::unit(5, "pt")
      ) +
      geom_text(
        aes(label = round(avaliacao, 1)),
        vjust = -0.4,
        color = "#0A1F44",
        fontface = "bold",
        size = 4.5
      ) +
      labs(
        x = NULL,
        y = paste0(input$indicador_escolhido, " (%)"),
        title = paste("Media da avaliacao:", input$indicador_escolhido)
      ) +
      theme_minimal(base_size = 13) +
      theme(
        panel.grid.major.x = element_blank(),
        panel.grid.minor = element_blank(),
        axis.text.x = element_text(face = "bold"),
        plot.title = element_text(face = "bold", hjust = 0.5, color = "#0A1F44")
      )
  })

  output$graf_correlacao_indicador <- renderPlot({
    req(input$tabs == "indicadores_tab")

    df_plot <- dados_indicador_detalhe()
    resumo_cor <- resumo_correlacao_indicador()
    cor_label <- resumo_cor$correlacao[[1]]

    ggplot(df_plot, aes(x = avaliacao_plot, y = produtividade_plot)) +
      geom_point(
        aes(color = tratamento),
        size = 3.2,
        alpha = 0.8
      ) +
      geom_smooth(
        method = "lm",
        se = TRUE,
        color = "#D4A017",
        linewidth = 1.2
      ) +
      labs(
        x = paste0(input$indicador_escolhido, " (%)"),
        y = "Produtividade (sc/ha)",
        title = paste("Produtividade x", input$indicador_escolhido),
        subtitle = paste("Correlacao estimada:", ifelse(is.na(cor_label), "-", round(cor_label, 2))),
        color = "Tratamento"
      ) +
      theme_minimal(base_size = 13) +
      theme(
        panel.grid.minor = element_blank(),
        plot.title = element_text(face = "bold", color = "#0A1F44"),
        plot.subtitle = element_text(color = "#444444"),
        legend.position = "bottom"
      )
  })

  output$graf_boxplots_indicadores <- renderPlot({
    req(input$tabs == "indicadores_tab")

    df_plot <- dados_indicador_detalhe() |>
      transmute(
        tratamento,
        produtividade = produtividade_plot,
        avaliacao = avaliacao_plot
      ) |>
      tidyr::pivot_longer(
        cols = c(produtividade, avaliacao),
        names_to = "metrica",
        values_to = "valor"
      ) |>
      mutate(
        metrica = factor(
          metrica,
          levels = c("produtividade", "avaliacao"),
          labels = c("Produtividade (sc/ha)", paste0(input$indicador_escolhido, " (%)"))
        ),
        tratamento = factor(tratamento, levels = unique(as.character(tratamento)))
      )

    ggplot(df_plot, aes(x = tratamento, y = valor, fill = tratamento)) +
      geom_boxplot(alpha = 0.85, outlier.alpha = 0.45, width = 0.68) +
      facet_wrap(~ metrica, scales = "free_y", ncol = 1) +
      labs(
        x = NULL,
        y = NULL,
        title = "Distribuicao por tratamento"
      ) +
      theme_minimal(base_size = 13) +
      theme(
        panel.grid.major.x = element_blank(),
        panel.grid.minor = element_blank(),
        plot.title = element_text(face = "bold", color = "#0A1F44", hjust = 0.5),
        axis.text.x = element_text(face = "bold"),
        legend.position = "none",
        strip.text = element_text(face = "bold")
      )
  })

  output$graf_trat_indicador <- renderPlot({
    req(input$tabs == "resultados_tab")
    df_plot <- dados_resultado()
    df_jitter <- base_filtrada_tratamentos() |>
      filter(
        !is.na(tratamento),
        !is.na(.data[[coluna_produtividade]])
      ) |>
      transmute(
        tratamento = trimws(as.character(tratamento)),
        produtividade = .data[[coluna_produtividade]]
      )

    validate(
      need(nrow(df_plot) > 0, "Sem dados para exibir no grafico.")
    )

    max_prod <- max(df_plot$produtividade + df_plot$se, na.rm = TRUE)
    max_var <- max(df_plot$avaliacao, na.rm = TRUE)

    validate(
      need(is.finite(max_prod) && max_prod > 0, "Escala invalida para produtividade."),
      need(is.finite(max_var) && max_var > 0, "Escala invalida para avaliacao.")
    )

    fator_escala <- max_prod / max_var

    df_plot <- df_plot |>
      mutate(
        var_plot = avaliacao * fator_escala,
        label_barra = ifelse(
          is.na(grupo) | grupo == "",
          as.character(round(produtividade, 1)),
          paste0(round(produtividade, 1), " ", grupo)
        ),
        label_linha = paste0(round(avaliacao, 1), "%"),
        cor_barra = case_when(
          grupo == "a" ~ "#0A1F44",
          grupo == "b" ~ "#163A5F",
          grupo == "c" ~ "#2A567A",
          grupo == "d" ~ "#4A789C",
          TRUE ~ "#163A5F"
        )
      )

    y_bar_min <- min(df_plot$produtividade - df_plot$se, na.rm = TRUE)
    y_bar_max <- max(df_plot$produtividade + df_plot$se, na.rm = TRUE)
    y_line_min <- min(df_plot$var_plot, na.rm = TRUE)
    y_line_max <- max(df_plot$var_plot, na.rm = TRUE)
    y_base_min <- min(y_bar_min, y_line_min)
    y_base_max <- max(y_bar_max, y_line_max)
    faixa_y <- y_base_max - y_base_min
    margem <- max(faixa_y * 0.12, 1)
    nudge_linha <- max(faixa_y * 0.04, 1.2)
    altura_jitter <- max(faixa_y * 0.015, 0.12)
    lim_inf <- max(0, y_base_min - margem)
    lim_sup <- max(y_bar_max, y_line_max + nudge_linha) + margem

    df_plot <- df_plot |>
      mutate(y_label_barra = lim_inf + (produtividade - lim_inf) / 2)

    ggplot(df_plot, aes(x = tratamento)) +
      ggchicklet::geom_chicklet(
        aes(y = produtividade),
        fill = "grey65",
        alpha = 0.16,
        width = 0.68,
        radius = grid::unit(5, "pt"),
        position = ggplot2::position_nudge(x = 0.03, y = 0.12)
      ) +
      ggchicklet::geom_chicklet(
        aes(y = produtividade, fill = cor_barra),
        color = "#08172F",
        size = 0.7,
        width = 0.65,
        radius = grid::unit(5, "pt")
      ) +
      geom_errorbar(
        aes(ymin = produtividade - se * 2, ymax = produtividade + se * 2),
        width = 0.25,
        linewidth = 0.35,
        colour = "#303030"
      ) +
      geom_jitter(
        data = df_jitter,
        aes(x = tratamento, y = produtividade),
        inherit.aes = FALSE,
        width = 0.12,
        height = altura_jitter,
        size = 2.3,
        alpha = 0.6,
        colour = "#D4A017"
      ) +
      geom_line(aes(y = var_plot, group = 1), linewidth = 2.2, colour = "#D4A017") +
      geom_point(aes(y = var_plot), size = 4, colour = "#D4A017") +
      geom_label(
        aes(y = var_plot, label = label_linha),
        fill = scales::alpha("#E8C85A", 0.85),
        color = "#1A1A1A",
        size = 5.2,
        fontface = "bold",
        linewidth = 0.35,
        label.r = grid::unit(0.18, "lines"),
        label.padding = grid::unit(0.18, "lines"),
        nudge_y = nudge_linha
      ) +
      geom_text(
        aes(y = y_label_barra, label = label_barra),
        vjust = 0.5,
        hjust = 0.5,
        colour = "white",
        fontface = "bold",
        size = 5.3
      ) +
      scale_y_continuous(
        name = "Produtividade (sc/ha)",
        sec.axis = sec_axis(~ . / fator_escala, name = paste0(input$indicador_escolhido, " (%)"))
      ) +
      coord_cartesian(ylim = c(lim_inf, lim_sup), clip = "off") +
      labs(
        x = NULL,
        caption = "* Medias seguidas pela mesma letra nao diferem pelo teste de Scott-Knott (5%)."
      ) +
      theme_minimal(base_size = 13) +
      theme(
        axis.text.x = element_text(angle = 0, hjust = 0.5, size = 14, face = "bold"),
        panel.grid.major.x = element_blank(),
        legend.position = "none",
        axis.title.y.left = element_text(face = "bold"),
        axis.title.y.right = element_text(face = "bold"),
        plot.caption = element_text(size = 15, face = "bold", color = "#444444", hjust = 0)
      ) +
      scale_fill_identity()
  })

  output$tabela_descricao_ensaio <- DT::renderDT({
    df <- descricao_ensaio()

    validate(
      need(nrow(df) > 0, "Sem descricao disponivel para este ensaio.")
    )

    DT::datatable(
      df,
      rownames = FALSE,
      escape = FALSE,
      class = "compact stripe hover",
      colnames = c("Tratamento", "Produto", "Dose", "Aplicacao"),
      options = list(
        dom = "t",
        pageLength = 20,
        ordering = FALSE,
        autoWidth = TRUE,
        scrollX = FALSE,
        columnDefs = list(list(className = "dt-center", targets = 0:3))
      )
    ) |>
      DT::formatStyle(columns = names(df), `font-size` = "14px", padding = "10px") |>
      DT::formatStyle("tratamento", fontWeight = "bold", color = "#0A1F44") |>
      DT::formatStyle("produto", backgroundColor = "#F7F9FC")
  })

  output$tabela_resumo <- DT::renderDT({
    if (nrow(resumo_base()) == 0) {
      return(
        DT::datatable(
          data.frame(Mensagem = "Sem dados para os filtros selecionados."),
          rownames = FALSE,
          options = list(dom = "t")
        )
      )
    }

    resumo_base() |>
      mutate(
        media_prod = round(media_prod, 2),
        sd_prod = round(sd_prod, 2),
        cv_prod = round(cv_prod, 2),
        ganho_vs_testemunha = round(ganho_vs_testemunha, 2)
      ) |>
      select(
        tratamento, produto, dose, aplicacao,
        media_prod, sd_prod, cv_prod, scott_knott, ganho_vs_testemunha
      ) |>
      DT::datatable(
        rownames = FALSE,
        options = list(pageLength = 15, scrollX = TRUE)
      )
  })

  output$baixar_relatorio_pdf <- downloadHandler(
    filename = function() {
      ensaio_nome <- if (!is.null(input$filtro_ensaio) && nzchar(input$filtro_ensaio)) {
        input$filtro_ensaio
      } else {
        "painel"
      }

      ensaio_nome <- gsub("[^A-Za-z0-9_-]+", "_", ensaio_nome)
      paste0("relatorio_", ensaio_nome, "_", Sys.Date(), ".pdf")
    },
    content = function(file) {
      base_atual <- isolate(base_filtrada())
      req(nrow(base_atual) > 0)

      resumo_atual <- isolate(resumo_base())
      indicador_nome <- isolate(input$indicador_escolhido)
      tratamentos_escolhidos <- isolate(input$filtro_tratamento)
      dados_resultado_atual <- tryCatch(isolate(dados_resultado()), error = function(e) NULL)
      dados_correlacao_atual <- tryCatch(isolate(dados_indicador_detalhe()), error = function(e) NULL)
      resumo_cor_atual <- tryCatch(isolate(resumo_correlacao_indicador()), error = function(e) NULL)
      descricao_atual <- tryCatch(isolate(descricao_ensaio()), error = function(e) NULL)

      filtros_linhas <- c(
        paste("Safra:", rotulo_filtro(isolate(input$filtro_safra))),
        paste("Cultura:", rotulo_filtro(isolate(input$filtro_cultura))),
        paste("Segmento:", rotulo_filtro(isolate(input$filtro_segmento))),
        paste("Ensaio:", rotulo_filtro(isolate(input$filtro_ensaio))),
        paste("Tratamentos:", rotulo_filtro(tratamentos_escolhidos)),
        paste("Avaliacao:", rotulo_filtro(indicador_nome, vazio = "Nao selecionada"))
      )

      kpis_linhas <- c(
        paste("Registros filtrados:", format(nrow(base_atual), big.mark = ".", scientific = FALSE)),
        paste("Tratamentos distintos:", format(dplyr::n_distinct(base_atual$tratamento, na.rm = TRUE), big.mark = ".", scientific = FALSE)),
        paste("Produtividade media:", round(mean(base_atual[[coluna_produtividade]], na.rm = TRUE), 2), "sc/ha"),
        paste("Produtividade maxima:", round(max(base_atual[[coluna_produtividade]], na.rm = TRUE), 2), "sc/ha"),
        paste("Produtividade minima:", round(min(base_atual[[coluna_produtividade]], na.rm = TRUE), 2), "sc/ha")
      )

      if (!is.null(resumo_cor_atual) && nrow(resumo_cor_atual) > 0) {
        kpis_linhas <- c(
          kpis_linhas,
          paste("Correlacao (r):", round(resumo_cor_atual$correlacao[[1]], 3)),
          paste("R2:", round(resumo_cor_atual$r2[[1]], 3)),
          paste("Observacoes validas:", resumo_cor_atual$observacoes[[1]])
        )
      }

      grDevices::pdf(file = file, width = 11.69, height = 8.27, paper = "special", onefile = TRUE)
      on.exit(grDevices::dev.off(), add = TRUE)

      adicionar_pagina_texto_pdf(
        titulo = "Relatorio tecnico do painel experimental",
        subtitulo = paste("Gerado em", format(Sys.time(), "%d/%m/%Y %H:%M")),
        linhas = c(
          filtros_linhas,
          "",
          "Resumo executivo",
          kpis_linhas
        ),
        familia = "sans",
        tamanho = 1
      )

      if (!is.null(dados_resultado_atual) && nrow(dados_resultado_atual) > 0) {
        print(plot_relatorio_produtividade(dados_resultado_atual, indicador_nome))
      }

      if (!is.null(dados_correlacao_atual) && nrow(dados_correlacao_atual) > 1) {
        print(plot_relatorio_correlacao(dados_correlacao_atual, indicador_nome))
      }

      if (!is.null(resumo_atual) && nrow(resumo_atual) > 0) {
        resumo_export <- resumo_atual |>
          mutate(
            media_prod = round(media_prod, 2),
            sd_prod = round(sd_prod, 2),
            cv_prod = round(cv_prod, 2),
            ganho_vs_testemunha = round(ganho_vs_testemunha, 2)
          ) |>
          select(
            tratamento, produto, dose, aplicacao,
            media_prod, sd_prod, cv_prod, scott_knott, ganho_vs_testemunha
          )

        adicionar_paginas_tabela_pdf(
          df = resumo_export,
          titulo = "Resumo tecnico por tratamento",
          subtitulo = quebrar_em_linhas(paste(filtros_linhas, collapse = " | "), largura = 110)
        )
      }

      if (!is.null(descricao_atual) && nrow(descricao_atual) > 0) {
        adicionar_paginas_tabela_pdf(
          df = descricao_atual,
          titulo = "Descricao dos tratamentos do ensaio",
          subtitulo = paste("Ensaio:", rotulo_filtro(isolate(input$filtro_ensaio)))
        )
      }
    }
  )
}

shiny::shinyApp(ui = ui, server = server)
