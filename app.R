library(shiny)
library(shinyjs)
library(openxlsx)
library(stringr)
library(dplyr)
library(bslib)
library(DT)

# --- 核心邏輯 ---
expand_ref_des <- function(input_str) {
  if (is.na(input_str) || input_str == "" || is.null(input_str)) return(list(val = "", changed = FALSE))
  original_text <- str_trim(as.character(input_str))
  clean_text <- str_replace_all(original_text, "[，；;]", ",") %>% 
    str_replace_all("\\s+", "") %>%
    str_replace("^[,]+", "") %>% 
    str_replace("[,]+$", "")
  parts <- str_split(clean_text, ",")[[1]]
  expanded <- c(); current_prefix <- "" 
  for (p in parts) {
    if (p == "") next
    match_range <- str_match(p, "^([A-Za-z]+)(\\d+)-([A-Za-z]+)?(\\d+)$")
    match_single <- str_match(p, "^([A-Za-z]+)(\\d+)$")
    match_num_only <- str_match(p, "^(\\d+)(?:-(\\d+))?$")
    if (!is.na(match_range[1, 1])) {
      current_prefix <- match_range[1, 2]; start_num <- as.integer(match_range[1, 3]); end_num <- as.integer(match_range[1, 5])
      expanded <- c(expanded, paste0(current_prefix, start_num:end_num))
    } else if (!is.na(match_single[1, 1])) {
      current_prefix <- match_single[1, 2]; expanded <- c(expanded, p)
    } else if (!is.na(match_num_only[1, 1]) && current_prefix != "") {
      start_num <- as.integer(match_num_only[1, 2]); end_num <- if (!is.na(match_num_only[1, 3])) as.integer(match_num_only[1, 3]) else start_num
      expanded <- c(expanded, paste0(current_prefix, start_num:end_num))
    } else { expanded <- c(expanded, p) }
  }
  final_str <- paste(expanded, collapse = ",")
  return(list(val = final_str, changed = (final_str != original_text)))
}

### UI
ui <- navbarPage(
  title = "DRAM插件位置轉換工具", 
  id = "main_navbar",
  theme = bs_theme(version = 5, bootswatch = "flatly"),
  header = tagList(
    useShinyjs(),
    tags$head(
      tags$style(HTML("
        body { background-color: #f8f9fa; }
        .nav-content { max-width: 800px; margin: auto; padding-top: 50px; }
        .card-box { background: white; padding: 40px; border-radius: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
        .changed-text { color: red !important; font-weight: bold; }
        .dt-container { background: white; padding: 20px; border-radius: 10px; }
        table.dataTable td { white-space: nowrap; vertical-align: middle; }
      "))
    )
  ),
  
  tabPanel("1. 上傳與轉換", value = "tab_upload",
           div(class = "nav-content",
               div(class = "card-box",
                   h3(style="margin-bottom:20px;", icon("upload"), strong("上傳原始檔案")),
                   fileInput("file", "請上傳從 WMS 下載的原始檔案(.xlsx)", accept = ".xlsx", width = "100%"),
                   br(),
                   # 改為點擊後彈出下載視窗，避免瀏覽器攔截或誤抓 HTML
                   actionButton("process", "執行轉換並查看預覽", 
                                class = "btn-outline-success btn-lg w-100", 
                                icon = icon("magic"))
               )
           )
  ),
  
  tabPanel("2. 預覽結果", value = "tab_preview",
           div(class = "container-fluid", style = "padding: 20px;",
               fluidRow(
                 column(width = 9, h2(strong("轉換結果預覽"))),
                 column(width = 3, uiOutput("download_ui")) # 將下載按鈕放在預覽頁面頂端
               ),
               hr(),
               div(class = "dt-container", DTOutput("preview_table"))
           )
  )
)

### Server
server <- function(input, output, session) {
  
  data_holder <- reactiveValues(df = NULL, original_name = "")
  
  observeEvent(input$process, {
    req(input$file)
    data_holder$original_name <- input$file$name
    
    # 執行數據處理
    df <- read.xlsx(input$file$datapath, startRow = 4)
    if (!("置件位置" %in% colnames(df))) {
      showModal(modalDialog(title = "格式錯誤", "找不到『置件位置』欄位。"))
      return(NULL)
    }
    
    results <- lapply(df$置件位置, expand_ref_des)
    df$New_Value <- sapply(results, function(x) x$val)
    df$Is_Changed <- sapply(results, function(x) x$changed)
    
    data_holder$df <- df
    
    # 直接跳轉分頁
    updateNavbarPage(session, "main_navbar", selected = "tab_preview")
    
    # 彈出成功提示，並提供下載按鈕
    showNotification("轉換完成！可以在預覽頁面頂端下載檔案。", type = "message")
  })
  
  # 動態生成下載按鈕（確保資料存在才出現，且能正確下載 xlsx）
  output$download_ui <- renderUI({
    req(data_holder$df)
    downloadButton("actual_download", "下載轉換結果", class = "btn-danger w-100")
  })
  
  # 預覽表格渲染
  output$preview_table <- renderDT({
    req(data_holder$df)
    display_df <- data_holder$df %>%
      select(子料號, Component品名, 原始內容 = 置件位置, 轉換後結果 = New_Value, Is_Changed)
    
    display_df$轉換後結果 <- ifelse(display_df$Is_Changed, 
                               paste0("<span class='changed-text'>", display_df$轉換後結果, "</span>"), 
                               display_df$轉換後結果)
    
    datatable(
      display_df %>% select(-Is_Changed),
      escape = FALSE,
      extensions = c('ColReorder', 'FixedColumns'),
      options = list(
        dom = 'ft', pageLength = -1, scrollX = TRUE, scrollY = "600px",
        autoWidth = FALSE,
        columnDefs = list(
          list(width = '500px', targets = 0), # 子料號 2.5 倍寬
          list(width = '100px', targets = 2)  # 原始內容縮窄
        ),
        colReorder = TRUE,
        fixedColumns = list(leftColumns = 1)
      ),
      rownames = FALSE
    )
  })
  
  # 真正的下載處理器
  output$actual_download <- downloadHandler(
    filename = function() { paste0("Fixed_", data_holder$original_name) },
    content = function(file) {
      req(data_holder$df)
      wb <- loadWorkbook(input$file$datapath)
      df <- data_holder$df
      col_names <- colnames(read.xlsx(input$file$datapath, startRow = 4))
      target_col_idx <- which(col_names == "置件位置")
      
      writeData(wb, sheet = 1, x = df$New_Value, startCol = target_col_idx, startRow = 5, colNames = FALSE)
      
      highlight_style <- createStyle(fgFill = "#FFFF00")
      changed_rows <- which(df$Is_Changed == TRUE) + 4
      if(length(changed_rows) > 0) {
        addStyle(wb, sheet = 1, style = highlight_style, rows = changed_rows, cols = target_col_idx, gridExpand = TRUE)
      }
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui, server)