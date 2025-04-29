library(shiny)
library(bs4Dash)

ui <- bs4DashPage(
  title = "786MIII Median & IQR Converter for Meta-Analysis",
  header = bs4DashNavbar(),
  sidebar = bs4DashSidebar(
    skin = "light",
    bs4SidebarMenu(
      bs4SidebarMenuItem("Single Conversion", tabName = "single", icon = icon("calculator")),
      bs4SidebarMenuItem("Batch Conversion", tabName = "batch", icon = icon("table"))
    )
  ),
  body = bs4DashBody(
    bs4TabItems(
      # Single conversion tab
      bs4TabItem(
        tabName = "single",
        fluidRow(
          bs4Card(
            title = "Single Conversion Input",
            width = 6,
            status = "primary",
            solidHeader = TRUE,
            numericInput("median", "Median", value = 0),
            numericInput("iqr", "Interquartile Range (IQR)", value = 1)
          ),
          bs4Card(
            title = "Conversion Output",
            width = 6,
            status = "success",
            solidHeader = TRUE,
            verbatimTextOutput("conversion_result")
          )
        )
      ),
      # Batch conversion tab
      bs4TabItem(
        tabName = "batch",
        fluidRow(
          bs4Card(
            title = "Upload CSV for Batch Conversion",
            width = 6,
            status = "primary",
            solidHeader = TRUE,
            fileInput("file", "Choose CSV File", accept = ".csv"),
            br(),
            downloadButton("download_sample", "Download Sample CSV")
          ),
          bs4Card(
            title = "Batch Conversion Output",
            width = 6,
            status = "success",
            solidHeader = TRUE,
            tableOutput("batch_table"),
            br(),
            downloadButton("download_data", "Download Converted CSV")
          )
        )
      )
    )
  ),
  controlbar = bs4DashControlbar(),
  footer = bs4DashFooter()
)

server <- function(input, output, session) {
  
  # Single conversion output
  output$conversion_result <- renderPrint({
    median_val <- input$median
    iqr_val <- input$iqr
    mean_val <- median_val
    sd_val <- iqr_val / 1.35
    
    cat("Converted Data:\n")
    cat("Mean:", mean_val, "\n")
    cat("Standard Deviation (approx):", sd_val, "\n")
  })
  
  # Reactive expression to process the uploaded CSV file
  batchData <- reactive({
    req(input$file)
    df <- read.csv(input$file$datapath)
    if(!all(c("median", "iqr") %in% names(df))){
      showNotification("CSV file must contain columns 'median' and 'iqr'", type = "error")
      return(NULL)
    }
    # Conversion: set mean as median and sd as iqr/1.35
    df$mean <- df$median
    df$sd <- df$iqr / 1.35
    df
  })
  
  # Display the converted batch data in a table
  output$batch_table <- renderTable({
    batchData()
  })
  
  # Download handler for the converted CSV file
  output$download_data <- downloadHandler(
    filename = function() {
      paste("converted_data_", Sys.Date(), ".csv", sep = "")
    },
    content = function(file) {
      req(batchData())
      write.csv(batchData(), file, row.names = FALSE)
    }
  )
  
  # Download handler for a sample CSV file
  output$download_sample <- downloadHandler(
    filename = function() {
      "sample_data.csv"
    },
    content = function(file) {
      sample_data <- data.frame(
        median = c(10, 20, 30),
        iqr = c(5, 7, 10)
      )
      write.csv(sample_data, file, row.names = FALSE)
    }
  )
}

shinyApp(ui, server)
