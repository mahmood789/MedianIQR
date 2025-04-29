# ---------------------------
# 1. Load Packages & Setup
# ---------------------------
# Ensure these packages are installed: install.packages(c("shiny", "bs4Dash", "DT", "readr", "readxl", "dplyr", "writexl", "metafor", "esc", "dmetar", "clipr", "zip"))
library(shiny)
library(bs4Dash)
library(DT)
library(readr)
library(readxl)
library(dplyr)
library(writexl)
library(metafor) # For escalc (binary data)
library(esc)     # For many effect size conversions
library(dmetar)  # For NNT, pool.groups, se.from.p
library(clipr)   # For copy to clipboard
library(zip)     # For downloading all samples

# ---------------------------
# 2. Helper Functions for Conversions
# ---------------------------

# --- Functions to estimate Mean/SD from Summaries ---

# Simple approximation: Mean = Median, SD = IQR / 1.35
convert_simple_iqr <- function(median, iqr, n) {
  if (anyNA(c(median, iqr, n)) || iqr <= 0 || n < 1) {
    return(list(Est.Mean = NA, Est.SD = NA, Effect.Size=NA, SE=NA, Var=NA, CI.Lower=NA, CI.Upper=NA, Weight=NA, Measure=NA, Result=NA, Method = "Simple Approx (Median/IQR)", Error = "Invalid input (median, iqr > 0, n >= 1 required)"))
  }
  est_mean <- median
  est_sd <- iqr / 1.35
  # Add NA placeholders for standard output structure
  return(list(Est.Mean = est_mean, Est.SD = est_sd, Effect.Size=NA, SE=NA, Var=NA, CI.Lower=NA, CI.Upper=NA, Weight=NA, Measure=NA, Result=NA, Method = "Simple Approx (Median/IQR)", Error = NA))
}

# Hozo et al. (2005) method: Estimate mean and SD from min, max, median, and n
convert_hozo <- function(min_val, max_val, median_val, n) {
  if (anyNA(c(min_val, max_val, median_val, n)) || n < 3 || min_val > median_val || median_val > max_val || min_val == max_val) {
    return(list(Est.Mean = NA, Est.SD = NA, Effect.Size=NA, SE=NA, Var=NA, CI.Lower=NA, CI.Upper=NA, Weight=NA, Measure=NA, Result=NA, Method = "Hozo et al. (2005)", Error = "Invalid input (n>=3, min<=median<=max, min!=max required)"))
  }
  # Estimate Mean (Equation 4 from Hozo et al.)
  est_mean <- (min_val + 2 * median_val + max_val) / 4
  
  # Estimate SD (Equation 12 for n > 25, Equation 16 for n <= 25)
  if (n > 25) {
    # est_sd <- (max_val - min_val) / 4 # Original paper Eq 12 approximation
    est_sd <- (max_val - min_val) / (2 * qnorm((n - 1) / n)) # Improved version based on range, Wan Eq 8
  } else {
    # Using Hozo's Eq 16 (more accurate for small n)
    est_sd <- (1/sqrt(12)) * sqrt( ((min_val - 2*median_val + max_val)^2)/4 + (max_val - min_val)^2 )
  }
  return(list(Est.Mean = est_mean, Est.SD = est_sd, Effect.Size=NA, SE=NA, Var=NA, CI.Lower=NA, CI.Upper=NA, Weight=NA, Measure=NA, Result=NA, Method = "Hozo et al. (2005)", Error = NA))
}


# Wan et al. (2014) method: Estimate mean and SD from q1, median, q3, n, and optionally min/max
convert_wan <- function(q1, median_val, q3, n, min_val = NA, max_val = NA) {
  if (anyNA(c(q1, median_val, q3, n)) || n < 3 || q1 > median_val || median_val > q3) {
    return(list(Est.Mean = NA, Est.SD = NA, Effect.Size=NA, SE=NA, Var=NA, CI.Lower=NA, CI.Upper=NA, Weight=NA, Measure=NA, Result=NA, Method = "Wan et al. (2014)", Error = "Invalid input (n>=3, q1<=median<=q3 required)"))
  }
  
  # Check if optional min/max are provided and valid
  use_min_max <- !is.na(min_val) && !is.na(max_val) && is.numeric(min_val) && is.numeric(max_val)
  if(use_min_max && (min_val > q1 || q3 > max_val || min_val == max_val)){
    use_min_max <- FALSE # Treat as invalid if inconsistent
    warning("Min/Max values provided for Wan method are inconsistent with quartiles or invalid; using Scenario S1.")
  }
  
  # Estimate Mean (Equation 7 if min/max available & valid, else Equation 1)
  if (use_min_max) {
    # Scenario S3: Using min, q1, median, q3, max
    est_mean <- (min_val + 2 * q1 + 2 * median_val + 2 * q3 + max_val) / 8
    method_used <- "Wan et al. (2014) - S3"
  } else {
    # Scenario S1: Using q1, median, q3 only
    est_mean <- (q1 + median_val + q3) / 3
    method_used <- "Wan et al. (2014) - S1"
  }
  
  # Estimate SD (Equation 15 if min/max available & valid, else Equation 11)
  # Using the improved estimators from Wan et al.
  if (use_min_max) {
    # Scenario S3 (Eq 15)
    est_sd <- (max_val - min_val) / (2 * qnorm((n - 1) / (n + 1))) # Uses range estimator similar to Hozo but different quantile approx
  } else {
    # Scenario S1 (Eq 11)
    est_sd <- (q3 - q1) / (2 * qnorm((0.75 * n - 0.125) / (n + 0.25))) # Uses IQR estimator
  }
  
  return(list(Est.Mean = est_mean, Est.SD = est_sd, Effect.Size=NA, SE=NA, Var=NA, CI.Lower=NA, CI.Upper=NA, Weight=NA, Measure=NA, Result=NA, Method = method_used, Error = NA))
}

# --- Functions for Effect Size Calculation/Conversion ---

# Wrapper for esc functions
run_esc_conversion <- function(fun_name, args_list) {
  args_list <- args_list[!sapply(args_list, function(x) is.null(x) || is.na(x))]
  valid_es_types <- c("d", "g", "r", "cox.or")
  if (!"es.type" %in% names(args_list) || !args_list$es.type %in% valid_es_types) {
    if (fun_name %in% c("mean_se", "B", "beta", "rpb", "f", "t")) {
      args_list$es.type <- "d"
    } else if (fun_name == "chisq") {
      args_list$es.type <- "cox.or"
    } else {
      args_list$es.type <- "d"
    }
  }
  
  result <- tryCatch({
    do.call(esc::esc, c(list(type = fun_name), args_list))
  }, error = function(e) list(error = e$message))
  
  if (!is.null(result$error)) {
    return(list(Est.Mean=NA, Est.SD=NA, Effect.Size = NA, SE = NA, Var = NA, CI.Lower = NA, CI.Upper = NA, Weight = NA, Measure = NA, Result=NA, Method = paste("esc::esc_", fun_name, sep=""), Error = result$error))
  } else {
    return(list(
      Est.Mean=NA, Est.SD=NA,
      Effect.Size = ifelse(!is.null(result$es), result$es, NA),
      SE = ifelse(!is.null(result$se), result$se, NA),
      Var = ifelse(!is.null(result$var), result$var, NA),
      CI.Lower = ifelse(!is.null(result$ci.lo), result$ci.lo, NA),
      CI.Upper = ifelse(!is.null(result$ci.hi), result$ci.hi, NA),
      Weight = ifelse(!is.null(result$w), result$w, NA),
      Measure = ifelse(!is.null(result$measure), result$measure, NA),
      Result = NA,
      Method = paste("esc::esc_", fun_name, sep=""),
      Error = NA))
  }
}

# Wrapper for dmetar functions (Corrected Error Handling and Output)
run_dmetar_conversion <- function(fun_name, args_list) {
  args_list <- args_list[!sapply(args_list, function(x) is.null(x) || is.na(x))]
  result_val <- tryCatch({
    do.call(fun_name, args_list)
  }, error = function(e) {
    # Return the error message itself if do.call fails
    e$message
  })
  
  # Check if the result is an error message (string)
  if (is.character(result_val) && length(result_val) == 1) {
    return(list(Est.Mean=NA, Est.SD=NA, Effect.Size = NA, SE = NA, Var = NA, CI.Lower = NA, CI.Upper = NA, Weight = NA, Measure = NA, Result = NA, Method = paste("dmetar::", fun_name, sep=""), Error = result_val))
  } else {
    # Process successful results
    res_text <- NULL
    if(fun_name == "pool.groups"){
      # Check if expected names exist and are numeric before rounding
      if(all(c("n.pooled", "m.pooled", "sd.pooled") %in% names(result_val)) &&
         is.numeric(result_val$n.pooled) && is.numeric(result_val$m.pooled) && is.numeric(result_val$sd.pooled)) {
        res_text <- paste("Pooled N:", result_val$n.pooled, "Pooled Mean:", round(result_val$m.pooled,3), "Pooled SD:", round(result_val$sd.pooled,3))
      } else {
        res_text <- "Pooling failed or returned unexpected structure."
      }
    } else if (fun_name == "NNT") {
      if(all(c("NNT", "CI.NNT") %in% names(result_val)) && is.numeric(result_val$NNT) && is.numeric(result_val$CI.NNT) && length(result_val$CI.NNT) == 2) {
        res_text <- paste("NNT:", round(result_val$NNT, 2), "(", round(result_val$CI.NNT[1], 2), "to", round(result_val$CI.NNT[2], 2), ")")
      } else {
        res_text <- "NNT calculation failed or returned unexpected structure."
      }
    } else if (fun_name == "se.from.p"){
      if(is.numeric(result_val) && length(result_val) == 1){
        res_text <- paste("Estimated SE:", round(result_val, 4))
      } else {
        res_text <- "SE estimation failed or returned unexpected structure."
      }
    } else {
      res_text <- paste(capture.output(result_val), collapse="\n") # Generic fallback
    }
    return(list(Est.Mean=NA, Est.SD=NA, Effect.Size = NA, SE = NA, Var = NA, CI.Lower = NA, CI.Upper = NA, Weight = NA, Measure = NA, Result = res_text, Method = paste("dmetar::", fun_name, sep=""), Error = NA))
  }
}


# Wrapper for metafor::escalc for binary 2x2 data
convert_2x2 <- function(ai, bi, ci, di, measure) {
  if(anyNA(c(ai, bi, ci, di))) {
    return(list(Est.Mean=NA, Est.SD=NA, Effect.Size = NA, SE=NA, Var = NA, CI.Lower = NA, CI.Upper = NA, Weight=NA, Measure=measure, Result=NA, Method="metafor::escalc", Error = "Missing values in 2x2 table input"))
  }
  if(any(c(ai, bi, ci, di) < 0)){
    return(list(Est.Mean=NA, Est.SD=NA, Effect.Size = NA, SE=NA, Var = NA, CI.Lower = NA, CI.Upper = NA, Measure=measure, Result=NA, Method="metafor::escalc", Error = "Counts (a,b,c,d) cannot be negative."))
  }
  
  dat <- data.frame(ai=ai, bi=bi, ci=ci, di=di)
  res <- tryCatch({
    metafor::escalc(measure=measure, ai=ai, bi=bi, ci=ci, di=di, data=dat, add=1/2, to="only0") # Add 1/2 continuity correction to zero cells only
  }, error = function(e) list(error=e$message))
  
  if(!is.null(res$error)) {
    return(list(Est.Mean=NA, Est.SD=NA, Effect.Size = NA, SE=NA, Var = NA, CI.Lower = NA, CI.Upper = NA, Weight=NA, Measure=measure, Result=NA, Method="metafor::escalc", Error = res$error))
  } else {
    z <- qnorm(0.975)
    se_calc <- sqrt(res$vi) # Calculate SE from variance
    ci.lb <- res$yi - z * se_calc
    ci.ub <- res$yi + z * se_calc
    if(measure %in% c("OR", "RR")){
      es_disp <- exp(res$yi)
      ci.lb_disp <- exp(ci.lb)
      ci.ub_disp <- exp(ci.ub)
    } else { # RD
      es_disp <- res$yi
      ci.lb_disp <- ci.lb
      ci.ub_disp <- ci.ub
    }
    return(list(Est.Mean=NA, Est.SD=NA, Effect.Size = es_disp, SE = se_calc, Var = res$vi, CI.Lower = ci.lb_disp, CI.Upper = ci.ub_disp, Weight=NA, Measure=measure, Result=NA, Method="metafor::escalc", Error = NA))
  }
}


#------------------------------------------------------------------
# Sample Data Generation (Expanded)
#------------------------------------------------------------------
generateSampleData <- function(convType) {
  # Ensure consistent column names across all samples
  base_df <- data.frame(study = character(), stringsAsFactors = FALSE)
  
  df <- switch(convType,
               "Simple Approx (Median/IQR)" = data.frame(study = c("A", "B"), median = c(10, 15), iqr = c(5, 8), n = c(50, 60)),
               "Hozo et al. (2005)" = data.frame(study = c("C", "D"), min_val = c(2, 5), median_val = c(8, 12), max_val = c(15, 25), n = c(30, 100)),
               "Wan et al. (2014)" = data.frame(study = c("E", "F", "G"), q1 = c(5, 10, 12), median_val = c(8, 15, 16), q3 = c(12, 22, 20), n = c(40, 80, 120), min_val = c(NA, 2, 5), max_val = c(NA, 30, 25)),
               "Mean & SE to d/g" = data.frame(study=c("H","I"), grp1m = 8.5, grp1se = 1.5, grp1n = 50, grp2m = 11, grp2se = 1.8, grp2n = 60, es.type = "d"),
               "Unstd Reg Coef (b) to d/r" = data.frame(study=c("J"), b = 3.3, sdy = 5, grp1n = 100, grp2n = 150, es.type = "d"),
               "Std Reg Coef (beta) to d/r" = data.frame(study=c("K"), beta = 0.32, sdy = 5, grp1n = 100, grp2n = 150, es.type = "d"),
               "Point-biserial r to d/r" = data.frame(study=c("L"), r = 0.25, grp1n = 99, grp2n = 120, es.type = "d"),
               "ANOVA F-value to d/g" = data.frame(study=c("M"), f = 5.04, grp1n = 519, grp2n = 528, es.type = "g"),
               "t-Test to d/g" = data.frame(study=c("N"), t = 3.3, grp1n = 100, grp2n = 150, es.type = "d"),
               "p-value to SE" = data.frame(study=c("O"), effect_size = 0.71, p = 0.013, N = 71, effect.size.type = "difference"),
               "Chi-squared to OR" = data.frame(study=c("P"), chisq = 7.9, totaln = 100, es.type = "cox.or"),
               "Pool Groups (Mean/SD)" = data.frame(study=c("Q"), n1 = 50, n2 = 50, m1 = 3.5, m2 = 4, sd1 = 3, sd2 = 3.8),
               "NNT from d" = data.frame(study=c("R"), d = 0.245, CER = 0.35),
               "2x2 Table to OR/RR/RD" = data.frame(study=c("S", "T"), ai=c(10, 25), bi=c(40, 35), ci=c(15, 30), di=c(35, 25), measure=c("OR", "RR")),
               data.frame(study=character()) # Default empty
  )
  # Ensure 'study' column exists if df is not empty
  if(nrow(df) > 0 && !"study" %in% names(df)){
    df$study <- paste("Row", 1:nrow(df))
  } else if (nrow(df) == 0) {
    df <- data.frame(study=character()) # Ensure it's a dataframe even if empty
  }
  return(df)
}


# List of all conversion types (grouped)
conversionMethodsList <- list(
  "Estimate Mean/SD" = c("Simple Approx (Median/IQR)", "Hozo et al. (2005)", "Wan et al. (2014)"),
  "Convert Stats to d/g/r" = c("Mean & SE to d/g",
                               "Unstd Reg Coef (b) to d/r",
                               "Std Reg Coef (beta) to d/r",
                               "Point-biserial r to d/r",
                               "ANOVA F-value to d/g",
                               "t-Test to d/g",
                               "Chi-squared to OR"), # OR is often converted later
  "Convert Binary (2x2) Data" = c("2x2 Table to OR/RR/RD"),
  "Other Conversions" = c("p-value to SE", "Pool Groups (Mean/SD)", "NNT from d")
)
allConvTypes <- unlist(conversionMethodsList, use.names = FALSE)

#------------------------------------------------------------------
# UI (Expanded)
#------------------------------------------------------------------
ui <- bs4DashPage(
  title = "MetaStatConverter: Estimate & Convert Effect Sizes",
  header = bs4DashNavbar(title = "MetaStatConverter"),
  sidebar = bs4DashSidebar(
    skin = "light",
    bs4SidebarMenu(
      id = "sidebarMenu",
      bs4SidebarMenuItem("Single Conversion", tabName = "single", icon = icon("calculator")),
      bs4SidebarMenuItem("Batch Conversion", tabName = "batch", icon = icon("table")),
      bs4SidebarMenuItem("History", tabName = "history", icon = icon("history")),
      bs4SidebarMenuItem("Instructions & About", tabName = "about", icon = icon("info-circle"))
    )
  ),
  body = bs4DashBody(
    # Removed the use_clipr() call from here
    bs4TabItems(
      # Single Conversion Tab
      bs4TabItem(
        tabName = "single",
        fluidRow(
          bs4Card(
            title = "Single Conversion Input",
            width = 6,
            status = "primary",
            solidHeader = TRUE,
            collapsible = TRUE,
            # Use grouped choices
            selectInput("singleMethod", "Select Conversion Type:", choices = conversionMethodsList),
            # Dynamic inputs based on method
            uiOutput("singleInputsUI"),
            br(),
            actionButton("calculateSingle", "Calculate", icon = icon("play-circle"), class = "btn-success")
          ),
          bs4Card(
            title = "Single Conversion Output",
            width = 6,
            status = "success",
            solidHeader = TRUE,
            collapsible = TRUE,
            verbatimTextOutput("singleResultOutput"),
            br(),
            uiOutput("copyButtonUI") # Placeholder for conditional copy button
          )
        )
      ),
      # Batch Conversion Tab
      bs4TabItem(
        tabName = "batch",
        fluidRow(
          bs4Card(
            title = "Batch Conversion Setup",
            width = 12,
            status = "primary",
            solidHeader = TRUE,
            collapsible = TRUE,
            selectInput("batchMethod", "Select Conversion Method for Batch File:", choices = conversionMethodsList),
            helpText("Ensure your CSV/Excel file has columns matching the required inputs (case-insensitive) for the selected method."),
            fileInput("batchFile", "Upload CSV or Excel File", accept = c(".csv", ".xlsx")),
            # Individual download buttons for each sample type
            h5("Download Sample CSV Templates:"),
            lapply(allConvTypes, function(conv) {
              downloadButton(paste0("dl_sample_", gsub("[^A-Za-z0-9]", "_", conv)),
                             paste("Sample:", conv),
                             class = "btn-sm btn-info", style="margin: 5px;")
            }),
            hr(),
            downloadButton("downloadAllSamples", "Download All Sample CSVs (ZIP)", class = "btn-warning"),
            hr(),
            actionButton("processBatch", "Process Batch File", icon = icon("cogs"), class = "btn-primary")
          )
        ),
        fluidRow(
          bs4Card(
            title = "Batch Conversion Results",
            width = 12,
            status = "success",
            solidHeader = TRUE,
            collapsible = TRUE,
            collapsed = FALSE, # Start expanded
            DTOutput("batchResultsTable"),
            br(),
            downloadButton("downloadBatchResults", "Download Results as CSV")
          )
        )
      ),
      # History Tab
      bs4TabItem(
        tabName = "history",
        fluidRow(
          bs4Card(
            title = "Single Conversion History",
            status = "secondary",
            width = 12,
            collapsible = TRUE,
            solidHeader = TRUE,
            DTOutput("historyTable"),
            br(),
            actionButton("clearHistory", "Clear History", icon = icon("trash"), class = "btn-danger")
          )
        )
      ),
      # Instructions & About Tab
      bs4TabItem(
        tabName = "about",
        fluidRow(
          bs4Card(
            title = "Instructions & About",
            status = "info",
            solidHeader = TRUE,
            width = 12,
            collapsible = FALSE,
            HTML("
              <h4>Purpose</h4>
              <p>This application estimates sample means and standard deviations (SD) from other summary statistics (like medians, quartiles, ranges) and converts various reported statistics (t-values, F-values, regression coefficients, 2x2 tables, etc.) into common effect size metrics (Cohen's d, Hedges' g, correlation r, Odds Ratio, Risk Ratio, Risk Difference) suitable for meta-analysis.</p>
              <h4>Methods Available</h4>
              <h5>Estimate Mean/SD:</h5>
              <ul>
                <li><b>Simple Approx (Median/IQR):</b> Mean ≈ Median, SD ≈ IQR / 1.35. Rough approximation. Requires: <code>median, iqr, n</code>.</li>
                <li><b>Hozo et al. (2005):</b> Uses min, max, median, n (n>=3). Assumes approx. normality for SD. Ref: BMC Med Res Methodol 5:13. Requires: <code>min_val, max_val, median_val, n</code>.</li>
                <li><b>Wan et al. (2014):</b> Uses Q1, median, Q3, n (n>=3). Optionally uses min, max for better accuracy (S3 vs S1). Ref: BMC Med Res Methodol 14:135. Requires: <code>q1, median_val, q3, n</code>. Optional: <code>min_val, max_val</code>.</li>
              </ul>
               <h5>Convert Stats to d/g/r:</h5>
              <ul>
                 <li><b>Mean & SE to d/g:</b> Uses <code>esc::esc_mean_se</code>. Requires: <code>grp1m, grp1se, grp1n, grp2m, grp2se, grp2n, es.type ('d' or 'g')</code>.</li>
                 <li><b>Unstd Reg Coef (b) to d/r:</b> Uses <code>esc::esc_B</code>. Requires: <code>b, sdy, grp1n, grp2n, es.type ('d' or 'r')</code>.</li>
                 <li><b>Std Reg Coef (beta) to d/r:</b> Uses <code>esc::esc_beta</code>. Requires: <code>beta, sdy, grp1n, grp2n, es.type ('d' or 'r')</code>.</li>
                 <li><b>Point-biserial r to d/r:</b> Uses <code>esc::esc_rpb</code>. Requires: <code>r, grp1n, grp2n, es.type ('d' or 'r')</code>.</li>
                 <li><b>ANOVA F-value to d/g:</b> Uses <code>esc::esc_f</code>. Requires: <code>f, grp1n, grp2n, es.type ('d' or 'g')</code>.</li>
                 <li><b>t-Test to d/g:</b> Uses <code>esc::esc_t</code>. Requires: <code>t, grp1n, grp2n, es.type ('d' or 'g')</code>.</li>
                 <li><b>Chi-squared to OR:</b> Uses <code>esc::esc_chisq</code> with <code>es.type='cox.or'</code>. Requires: <code>chisq, totaln</code>.</li>
              </ul>
               <h5>Convert Binary (2x2) Data:</h5>
              <ul>
                 <li><b>2x2 Table to OR/RR/RD:</b> Uses <code>metafor::escalc</code>. Requires: <code>ai, bi, ci, di</code> (events/non-events in group 1 & 2). Select desired measure (OR, RR, RD) in UI. Applies continuity correction for zero cells.</li>
              </ul>
              <h5>Other Conversions:</h5>
              <ul>
                 <li><b>p-value to SE:</b> Uses <code>dmetar::se.from.p</code>. Requires: <code>effect_size, p, N, effect.size.type ('difference' or 'ratio')</code>.</li>
                 <li><b>Pool Groups (Mean/SD):</b> Uses <code>dmetar::pool.groups</code>. Requires: <code>n1, m1, sd1, n2, m2, sd2</code>.</li>
                 <li><b>NNT from d:</b> Uses <code>dmetar::NNT</code>. Requires: <code>d</code>. Optional: <code>CER</code> (Control Event Rate).</li>
              </ul>
              <h4>Usage</h4>
              <p><b>Single Conversion:</b> Select a method group, then the specific conversion. Enter required values, click 'Calculate'. Result appears on the right. Use the 'Copy' button if needed.</p>
              <p><b>Batch Conversion:</b> Select the method. Upload a CSV/Excel file with matching column headers (case-insensitive). Download specific samples for guidance. Click 'Process'. Results table appears below, including new columns for estimated/converted values.</p>
              <h4>Limitations</h4>
              <p>Mean/SD estimations are approximate; accuracy depends on data distribution and method assumptions (see original papers). Effect size conversions rely on the validity of the input statistics. Batch processing requires exact column name matches (case-insensitive) for the selected method.</p>
              <p><b>References:</b> See original papers for Hozo et al. (2005) and Wan et al. (2014). Conversions primarily use the 'esc', 'dmetar', and 'metafor' R packages.</p>
            ")
          )
        )
      )
    )
  ),
  controlbar = bs4DashControlbar(),
  footer = bs4DashFooter()
)

#------------------------------------------------------------------
# Server
#------------------------------------------------------------------
server <- function(input, output, session) {
  
  # --- Reactive Value for History ---
  convHistory <- reactiveVal(data.frame(
    Time = character(),
    ConversionType = character(),
    Inputs = character(),
    Result = character(),
    stringsAsFactors = FALSE
  ))
  
  # --- Single Conversion ---
  
  # Dynamic UI for single conversion inputs
  output$singleInputsUI <- renderUI({
    method <- input$singleMethod
    ns <- session$ns # Get namespace for inputs
    
    # Define common numericInput parameters
    numInput <- function(id, label, value, min_val = NA, max_val = NA, step_val = NA) {
      numericInput(ns(id), label, value = value, min = min_val, max = max_val, step = step_val)
    }
    selInput <- function(id, label, choices, selected) {
      selectInput(ns(id), label, choices = choices, selected = selected)
    }
    
    # Generate UI based on selected method
    switch(method,
           "Simple Approx (Median/IQR)" = tagList(numInput("s_median", "Median:", 10), numInput("s_iqr", "IQR:", 5, min_val = 0.001), numInput("s_n_simple", "Sample Size (n):", 50, min_val = 1)), # IQR > 0
           "Hozo et al. (2005)" = tagList(numInput("s_min_hozo", "Minimum Value:", 0), numInput("s_max_hozo", "Maximum Value:", 20), numInput("s_median_hozo", "Median:", 8), numInput("s_n_hozo", "Sample Size (n):", 30, min_val = 3)),
           "Wan et al. (2014)" = tagList(numInput("s_q1_wan", "Q1:", 5), numInput("s_median_wan", "Median:", 10), numInput("s_q3_wan", "Q3:", 15), numInput("s_n_wan", "Sample Size (n):", 40, min_val = 3), checkboxInput(ns("s_use_minmax_wan"), "Use Min/Max?", FALSE), conditionalPanel(condition = "input.s_use_minmax_wan == true", ns=ns, numInput("s_min_wan", "Min (Optional):", NA), numInput("s_max_wan", "Max (Optional):", NA))),
           "Mean & SE to d/g" = tagList(numInput("grp1m", "Group 1 Mean:", 8.5), numInput("grp1se", "Group 1 SE:", 1.5, min_val=0), numInput("grp1n", "Group 1 n:", 50, min_val = 3), numInput("grp2m", "Group 2 Mean:", 11), numInput("grp2se", "Group 2 SE:", 1.8, min_val=0), numInput("grp2n", "Group 2 n:", 60, min_val = 3), selInput("es.type", "Effect Size Type:", c("d", "g"), "d")),
           "Unstd Reg Coef (b) to d/r" = tagList(numInput("b", "Coefficient (b):", 3.3), numInput("sdy", "SD of y:", 5, min_val=0.001), numInput("grp1n", "Group 1 n:", 100, min_val = 3), numInput("grp2n", "Group 2 n:", 150, min_val = 3), selInput("es.type", "Effect Size Type:", c("d", "r"), "d")),
           "Std Reg Coef (beta) to d/r" = tagList(numInput("beta", "Standardized Beta:", 0.32), numInput("sdy", "SD of y:", 5, min_val=0.001), numInput("grp1n", "Group 1 n:", 100, min_val = 3), numInput("grp2n", "Group 2 n:", 150, min_val = 3), selInput("es.type", "Effect Size Type:", c("d", "r"), "d")),
           "Point-biserial r to d/r" = tagList(numInput("r", "Point-Biserial r:", 0.25, min=-1, max=1), numInput("grp1n", "Group 1 n:", 99, min_val = 3), numInput("grp2n", "Group 2 n:", 120, min_val = 3), selInput("es.type", "Effect Size Type:", c("d", "r"), "d")),
           "ANOVA F-value to d/g" = tagList(numInput("f", "F-value:", 5.04, min=0), numInput("grp1n", "Group 1 n:", 519, min_val = 3), numInput("grp2n", "Group 2 n:", 528, min_val = 3), selInput("es.type", "Effect Size Type:", c("d", "g"), "g")),
           "t-Test to d/g" = tagList(numInput("t", "t-value:", 3.3), numInput("grp1n", "Group 1 n:", 100, min_val = 3), numInput("grp2n", "Group 2 n:", 150, min_val = 3), selInput("es.type", "Effect Size Type:", c("d", "g"), "d")),
           "p-value to SE" = tagList(numInput("effect_size", "Effect Size:", 0.71), numInput("p", "p-value:", 0.013, min = 0, max = 1), numInput("N", "Total N:", 71, min = 3), selInput("effect.size.type", "Effect Size Type:", c("difference", "ratio"), "difference")),
           "Chi-squared to OR" = tagList(numInput("chisq", "Chi-squared value:", 7.9, min=0), numInput("totaln", "Total N:", 100, min = 1), selInput("es.type", "Effect Size Type:", c("cox.or"), "cox.or")), # Only cox.or supported by esc_chisq
           "Pool Groups (Mean/SD)" = tagList(numInput("n1", "Group 1 n:", 50, min = 1), numInput("m1", "Group 1 Mean:", 3.5), numInput("sd1", "Group 1 SD:", 3, min=0), numInput("n2", "Group 2 n:", 50, min = 1), numInput("m2", "Group 2 Mean:", 4), numInput("sd2", "Group 2 SD:", 3.8, min=0)),
           "NNT from d" = tagList(numInput("d", "Effect Size (d):", 0.245), checkboxInput(ns("nnt_useCER"), "Specify CER?", FALSE), conditionalPanel(condition = "input.nnt_useCER == true", ns=ns, numInput("nnt_CER", "Control Event Rate (CER):", 0.35, min = 0, max = 1))),
           "2x2 Table to OR/RR/RD" = tagList(numericInput(ns("ai"), "Group 1 Events (a):", 10, min=0), numericInput(ns("bi"), "Group 1 Non-Events (b):", 40, min=0), numericInput(ns("ci"), "Group 2 Events (c):", 15, min=0), numericInput(ns("di"), "Group 2 Non-Events (d):", 35, min=0), selInput("measure_2x2", "Desired Measure:", c("OR", "RR", "RD"), "OR"))
    )
  })
  
  # Calculate single conversion result
  singleResult <- eventReactive(input$calculateSingle, {
    method <- input$singleMethod
    res <- NULL
    error_msg <- NULL
    inputs_list <- list() # Store inputs for history
    
    # --- Estimate Mean/SD ---
    if (method == "Simple Approx (Median/IQR)") {
      inputs_list <- list(median=input$s_median, iqr=input$s_iqr, n=input$s_n_simple)
      if(anyNA(inputs_list[c("median","iqr","n")]) || input$s_iqr <= 0 || input$s_n_simple < 1) error_msg <- "Invalid input for Simple Approx."
      else res <- convert_simple_iqr(input$s_median, input$s_iqr, input$s_n_simple)
    } else if (method == "Hozo et al. (2005)") {
      inputs_list <- list(min_val=input$s_min_hozo, max_val=input$s_max_hozo, median_val=input$s_median_hozo, n=input$s_n_hozo)
      if(anyNA(inputs_list) || input$s_n_hozo < 3) error_msg <- "Hozo requires n>=3 and non-NA min, max, median."
      else if(input$s_min_hozo > input$s_median_hozo || input$s_median_hozo > input$s_max_hozo) error_msg <- "Values must be: Min <= Median <= Max."
      else if(input$s_min_hozo == input$s_max_hozo) error_msg <- "Min and Max cannot be equal."
      else res <- convert_hozo(input$s_min_hozo, input$s_max_hozo, input$s_median_hozo, input$s_n_hozo)
    } else if (method == "Wan et al. (2014)") {
      inputs_list <- list(q1=input$s_q1_wan, median_val=input$s_median_wan, q3=input$s_q3_wan, n=input$s_n_wan, use_minmax=input$s_use_minmax_wan, min_val=input$s_min_wan, max_val=input$s_max_wan)
      if(anyNA(inputs_list[c("q1","median_val","q3","n")]) || input$s_n_wan < 3) error_msg <- "Wan requires n>=3 and non-NA Q1, Median, Q3."
      else if(input$s_q1_wan > input$s_median_wan || input$s_median_wan > input$s_q3_wan) error_msg <- "Quartiles must be: Q1 <= Median <= Q3."
      else {
        min_v <- if(input$s_use_minmax_wan && !is.na(input$s_min_wan)) input$s_min_wan else NA
        max_v <- if(input$s_use_minmax_wan && !is.na(input$s_max_wan)) input$s_max_wan else NA
        if(input$s_use_minmax_wan && (is.na(min_v) || is.na(max_v))) { error_msg <- "Min/Max required if checkbox checked." }
        else if (input$s_use_minmax_wan && (min_v > input$s_q1_wan || max_v < input$s_q3_wan)) { error_msg <- "Min must be <= Q1 and Max >= Q3." }
        else { res <- convert_wan(input$s_q1_wan, input$s_median_wan, input$s_q3_wan, input$s_n_wan, min_v, max_v) }
      }
    }
    # --- Convert Stats to d/g/r ---
    else if (method == "Mean & SE to d/g") {
      # Need to access the correct input ID (ns'd version)
      inputs_list <- list(grp1m=input$grp1m, grp1se=input$grp1se, grp1n=input$grp1n, grp2m=input$grp2m, grp2se=input$grp2se, grp2n=input$grp2n, es.type=input$es.type)
      if(anyNA(inputs_list[c("grp1m","grp1se","grp1n","grp2m","grp2se","grp2n")]) || input$grp1se <= 0 || input$grp2se <= 0 || input$grp1n < 3 || input$grp2n < 3) error_msg <- "Invalid input for Mean/SE conversion."
      else res <- run_esc_conversion("mean_se", inputs_list)
    } else if (method == "Unstd Reg Coef (b) to d/r") {
      inputs_list <- list(b=input$b, sdy=input$sdy, grp1n=input$grp1n, grp2n=input$grp2n, es.type=input$es.type)
      if(anyNA(inputs_list[c("b","sdy","grp1n","grp2n")]) || input$sdy <= 0 || input$grp1n < 3 || input$grp2n < 3) error_msg <- "Invalid input for Unstd Reg Coef conversion."
      else res <- run_esc_conversion("B", inputs_list)
    } else if (method == "Std Reg Coef (beta) to d/r") {
      inputs_list <- list(beta=input$beta, sdy=input$sdy, grp1n=input$grp1n, grp2n=input$grp2n, es.type=input$es.type)
      if(anyNA(inputs_list[c("beta","sdy","grp1n","grp2n")]) || input$sdy <= 0 || input$grp1n < 3 || input$grp2n < 3) error_msg <- "Invalid input for Std Reg Coef conversion."
      else res <- run_esc_conversion("beta", inputs_list)
    } else if (method == "Point-biserial r to d/r") {
      inputs_list <- list(r=input$r, grp1n=input$grp1n, grp2n=input$grp2n, es.type=input$es.type)
      if(anyNA(inputs_list[c("r","grp1n","grp2n")]) || abs(input$r) >= 1 || input$grp1n < 3 || input$grp2n < 3) error_msg <- "Invalid input for Point-biserial r conversion."
      else res <- run_esc_conversion("rpb", inputs_list)
    } else if (method == "ANOVA F-value to d/g") {
      inputs_list <- list(f=input$f, grp1n=input$grp1n, grp2n=input$grp2n, es.type=input$es.type)
      if(anyNA(inputs_list[c("f","grp1n","grp2n")]) || input$f < 0 || input$grp1n < 3 || input$grp2n < 3) error_msg <- "Invalid input for ANOVA F conversion."
      else res <- run_esc_conversion("f", inputs_list)
    } else if (method == "t-Test to d/g") {
      inputs_list <- list(t=input$t, grp1n=input$grp1n, grp2n=input$grp2n, es.type=input$es.type)
      if(anyNA(inputs_list[c("t","grp1n","grp2n")]) || input$grp1n < 3 || input$grp2n < 3) error_msg <- "Invalid input for t-Test conversion."
      else res <- run_esc_conversion("t", inputs_list)
    } else if (method == "Chi-squared to OR") {
      inputs_list <- list(chisq=input$chisq, totaln=input$totaln, es.type=input$es.type)
      if(anyNA(inputs_list[c("chisq","totaln")]) || input$chisq < 0 || input$totaln < 1) error_msg <- "Invalid input for Chi-squared conversion."
      else res <- run_esc_conversion("chisq", inputs_list)
    }
    # --- Convert Binary (2x2) Data ---
    else if (method == "2x2 Table to OR/RR/RD") {
      inputs_list <- list(ai=input$ai, bi=input$bi, ci=input$ci, di=input$di, measure=input$measure_2x2)
      if(anyNA(inputs_list[c("ai","bi","ci","di")]) || any(c(input$ai, input$bi, input$ci, input$di) < 0)) error_msg <- "Invalid input for 2x2 table (must be non-negative integers)."
      else res <- convert_2x2(input$ai, input$bi, input$ci, input$di, input$measure_2x2)
    }
    # --- Other Conversions ---
    else if (method == "p-value to SE") {
      inputs_list <- list(effect_size=input$effect_size, p=input$p, N=input$N, effect.size.type=input$effect.size.type)
      if(anyNA(inputs_list[c("effect_size","p","N")]) || input$p < 0 || input$p > 1 || input$N < 3) error_msg <- "Invalid input for p-value conversion."
      else res <- run_dmetar_conversion("se.from.p", inputs_list)
    } else if (method == "Pool Groups (Mean/SD)") {
      inputs_list <- list(n1=input$n1, m1=input$m1, sd1=input$sd1, n2=input$n2, m2=input$m2, sd2=input$sd2)
      if(anyNA(inputs_list) || input$n1 < 1 || input$n2 < 1 || input$sd1 < 0 || input$sd2 < 0) error_msg <- "Invalid input for pooling groups."
      else res <- run_dmetar_conversion("pool.groups", inputs_list)
    } else if (method == "NNT from d") {
      cer_val <- if(input$nnt_useCER && !is.na(input$nnt_CER)) input$nnt_CER else NULL
      inputs_list <- list(d=input$d, CER=cer_val)
      if(is.na(input$d)) error_msg <- "Effect size d is required for NNT."
      else if(input$nnt_useCER && is.na(cer_val)) error_msg <- "CER must be provided if checkbox is checked."
      else if(!is.null(cer_val) && (cer_val < 0 || cer_val > 1)) error_msg <- "CER must be between 0 and 1."
      else res <- run_dmetar_conversion("NNT", inputs_list)
    }
    
    # Handle errors and update history
    final_result <- NULL
    if(!is.null(error_msg)){
      showNotification(error_msg, type = "error")
      final_result <- list(Error = error_msg) # Store error message
    } else if (!is.null(res$Error) && !is.na(res$Error)) {
      # Error message already shown by helper function if applicable
      final_result <- res # Store result containing error
    } else {
      final_result <- res # Store successful result
    }
    
    # Append conversion to history (only if calculation was attempted)
    isolate({
      history <- convHistory()
      # Create a string summary of inputs, handling NULLs gracefully
      input_summary <- paste(names(inputs_list), "=", sapply(inputs_list, function(x) {
        if(is.null(x)) "NULL" else if(is.na(x)) "NA" else as.character(x)
      }), collapse="; ")
      
      
      # Create a string summary of results or error
      result_summary <- if (!is.null(final_result$Error) && !is.na(final_result$Error)) {
        paste("Error:", final_result$Error)
      } else if (method %in% c("Simple Approx (Median/IQR)", "Hozo et al. (2005)", "Wan et al. (2014)")) {
        paste("Est.Mean=", round(final_result$Est.Mean, 3), "; Est.SD=", round(final_result$Est.SD, 3), "; Method=", final_result$Method)
      } else if (method %in% c("p-value to SE", "Pool Groups (Mean/SD)", "NNT from d")) {
        final_result$Result
      } else if (method == "2x2 Table to OR/RR/RD"){
        paste(final_result$Measure,"=", round(final_result$Effect.Size, 3), "; Var=", round(final_result$Var, 4), "; CI=[", round(final_result$CI.Lower, 3), ",", round(final_result$CI.Upper, 3), "]")
      } else { # Assume esc conversion result
        paste(final_result$Measure,"=", round(final_result$Effect.Size, 3), "; SE=", round(final_result$SE, 4), "; Var=", round(final_result$Var, 4), "; CI=[", round(final_result$CI.Lower, 3), ",", round(final_result$CI.Upper, 3), "]")
      }
      
      newRow <- data.frame(
        Time = format(Sys.time(), "%Y-%m-%d %H:%M:%S"),
        ConversionType = input$singleMethod,
        Inputs = input_summary,
        Result = result_summary,
        stringsAsFactors = FALSE
      )
      convHistory(rbind(history, newRow))
    })
    
    final_result # Return the result list (potentially with error)
  })
  
  # Display single conversion result
  output$singleResultOutput <- renderPrint({
    res <- singleResult()
    req(res) # Ensure calculation has run
    
    if (!is.null(res$Error) && !is.na(res$Error)) {
      cat("Error:", res$Error)
    } else {
      method <- isolate(input$singleMethod) # Get method again for formatting
      if (method %in% c("Simple Approx (Median/IQR)", "Hozo et al. (2005)", "Wan et al. (2014)")) {
        cat("Method Used:", res$Method, "\n")
        cat("Estimated Mean:", round(res$Est.Mean, 3), "\n")
        cat("Estimated SD:", round(res$Est.SD, 3), "\n")
      } else if (method %in% c("p-value to SE", "Pool Groups (Mean/SD)", "NNT from d")) {
        cat(res$Result)
      } else if (method == "2x2 Table to OR/RR/RD"){
        cat("Measure:", res$Measure, "\n")
        cat("Effect Size:", round(res$Effect.Size, 3), "\n")
        cat("Variance:", round(res$Var, 4), "\n")
        cat("95% CI:", paste0("[", round(res$CI.Lower, 3), ", ", round(res$CI.Upper, 3), "]"))
      } else { # Assume esc conversion result
        cat("Measure:", res$Measure, "\n")
        cat("Effect Size:", round(res$Effect.Size, 3), "\n")
        cat("Standard Error:", round(res$SE, 4), "\n")
        cat("Variance:", round(res$Var, 4), "\n")
        cat("95% CI:", paste0("[", round(res$CI.Lower, 3), ", ", round(res$CI.Upper, 3), "]"), "\n")
        cat("Weight:", round(res$Weight, 3), "\n")
      }
    }
  })
  
  # Render Copy button only if there's a valid result
  output$copyButtonUI <- renderUI({
    res <- singleResult()
    # Check if res exists and doesn't have an error before showing button
    if (!is.null(res) && (is.null(res$Error) || is.na(res$Error))) {
      actionButton(session$ns("copySingleResult"), "Copy Output to Clipboard", icon = icon("copy"), class = "btn-secondary")
    } else {
      NULL # Don't show button if no result or error
    }
  })
  
  
  # Copy single result to clipboard
  observeEvent(input$copySingleResult, {
    res <- singleResult()
    # Double check validity before copying
    if (!is.null(res) && (is.null(res$Error) || is.na(res$Error))) {
      method <- isolate(input$singleMethod)
      output_text <- capture.output( # Capture the exact text printed
        if (method %in% c("Simple Approx (Median/IQR)", "Hozo et al. (2005)", "Wan et al. (2014)")) {
          cat("Method Used:", res$Method, "\n")
          cat("Estimated Mean:", round(res$Est.Mean, 3), "\n")
          cat("Estimated SD:", round(res$Est.SD, 3), "\n")
        } else if (method %in% c("p-value to SE", "Pool Groups (Mean/SD)", "NNT from d")) {
          cat(res$Result)
        } else if (method == "2x2 Table to OR/RR/RD"){
          cat("Measure:", res$Measure, "\n")
          cat("Effect Size:", round(res$Effect.Size, 3), "\n")
          cat("Variance:", round(res$Var, 4), "\n")
          cat("95% CI:", paste0("[", round(res$CI.Lower, 3), ", ", round(res$CI.Upper, 3), "]"))
        } else { # Assume esc conversion result
          cat("Measure:", res$Measure, "\n")
          cat("Effect Size:", round(res$Effect.Size, 3), "\n")
          cat("Standard Error:", round(res$SE, 4), "\n")
          cat("Variance:", round(res$Var, 4), "\n")
          cat("95% CI:", paste0("[", round(res$CI.Lower, 3), ", ", round(res$CI.Upper, 3), "]"), "\n")
          cat("Weight:", round(res$Weight, 3), "\n")
        }
      )
      clipr::write_clip(paste(output_text, collapse="\n"))
      showNotification("Result copied to clipboard!", type = "message")
    } else {
      showNotification("Cannot copy - no valid result available.", type = "warning")
    }
  })
  
  # --- Batch Conversion ---
  
  # Reactive for uploaded batch data
  batchUploadData <- reactive({
    req(input$batchFile)
    f <- input$batchFile$datapath
    ext <- tools::file_ext(f)
    df <- tryCatch({
      if (ext == "csv") {
        read_csv(f, show_col_types = FALSE)
      } else if (ext %in% c("xls", "xlsx")) {
        readxl::read_excel(f)
      } else {
        stop("Unsupported file type. Please upload CSV or Excel.")
      }
    }, error = function(e) {
      showNotification(paste("Error reading file:", e$message), type = "error")
      NULL
    })
    df
  })
  
  # Process batch data
  batchProcessedData <- eventReactive(input$processBatch, {
    df_orig <- batchUploadData()
    req(df_orig)
    method <- input$batchMethod
    results_list <- list() # Store list of results per row
    
    # Define required columns based on method
    required_cols <- switch(method,
                            "Simple Approx (Median/IQR)" = c("median", "iqr", "n"),
                            "Hozo et al. (2005)" = c("min_val", "max_val", "median_val", "n"),
                            "Wan et al. (2014)" = c("q1", "median_val", "q3", "n"), # Optional: min_val, max_val
                            "Mean & SE to d/g" = c("grp1m", "grp1se", "grp1n", "grp2m", "grp2se", "grp2n", "es.type"),
                            "Unstd Reg Coef (b) to d/r" = c("b", "sdy", "grp1n", "grp2n", "es.type"),
                            "Std Reg Coef (beta) to d/r" = c("beta", "sdy", "grp1n", "grp2n", "es.type"),
                            "Point-biserial r to d/r" = c("r", "grp1n", "grp2n", "es.type"),
                            "ANOVA F-value to d/g" = c("f", "grp1n", "grp2n", "es.type"),
                            "t-Test to d/g" = c("t", "grp1n", "grp2n", "es.type"),
                            "p-value to SE" = c("effect_size", "p", "N", "effect.size.type"),
                            "Chi-squared to OR" = c("chisq", "totaln", "es.type"),
                            "Pool Groups (Mean/SD)" = c("n1", "m1", "sd1", "n2", "m2", "sd2"),
                            "NNT from d" = c("d"), # Optional: CER
                            "2x2 Table to OR/RR/RD" = c("ai", "bi", "ci", "di", "measure"),
                            c() # Default empty
    )
    
    # Check required columns (case-insensitive)
    df_lower_names <- tolower(names(df_orig))
    missing_cols <- setdiff(required_cols, df_lower_names)
    if (length(missing_cols) > 0) {
      showNotification(paste("Missing required columns for", method, ":", paste(missing_cols, collapse=", ")), type="error", duration=15)
      return(NULL)
    }
    # Rename to lowercase for processing
    df <- df_orig %>% rename_all(tolower)
    
    # Apply conversion row-wise
    for (i in 1:nrow(df)) {
      row <- df[i, ]
      res <- NULL
      args_list <- as.list(row)
      
      # --- Call appropriate function ---
      if (method == "Simple Approx (Median/IQR)") { res <- convert_simple_iqr(row$median, row$iqr, row$n) }
      else if (method == "Hozo et al. (2005)") { res <- convert_hozo(row$min_val, row$max_val, row$median_val, row$n) }
      else if (method == "Wan et al. (2014)") {
        min_v <- if("min_val" %in% names(row) && !is.na(row$min_val)) row$min_val else NA
        max_v <- if("max_val" %in% names(row) && !is.na(row$max_val)) row$max_val else NA
        res <- convert_wan(row$q1, row$median_val, row$q3, row$n, min_v, max_v)
      }
      else if (method == "Mean & SE to d/g") { res <- run_esc_conversion("mean_se", args_list[c("grp1m", "grp1se", "grp1n", "grp2m", "grp2se", "grp2n", "es.type")]) }
      else if (method == "Unstd Reg Coef (b) to d/r") { res <- run_esc_conversion("B", args_list[c("b", "sdy", "grp1n", "grp2n", "es.type")]) }
      else if (method == "Std Reg Coef (beta) to d/r") { res <- run_esc_conversion("beta", args_list[c("beta", "sdy", "grp1n", "grp2n", "es.type")]) }
      else if (method == "Point-biserial r to d/r") { res <- run_esc_conversion("rpb", args_list[c("r", "grp1n", "grp2n", "es.type")]) }
      else if (method == "ANOVA F-value to d/g") { res <- run_esc_conversion("f", args_list[c("f", "grp1n", "grp2n", "es.type")]) }
      else if (method == "t-Test to d/g") { res <- run_esc_conversion("t", args_list[c("t", "grp1n", "grp2n", "es.type")]) }
      else if (method == "Chi-squared to OR") { res <- run_esc_conversion("chisq", args_list[c("chisq", "totaln", "es.type")]) }
      else if (method == "2x2 Table to OR/RR/RD") { res <- convert_2x2(row$ai, row$bi, row$ci, row$di, row$measure) }
      else if (method == "p-value to SE") { res <- run_dmetar_conversion("se.from.p", args_list[c("effect_size", "p", "N", "effect.size.type")]) }
      else if (method == "Pool Groups (Mean/SD)") { res <- run_dmetar_conversion("pool.groups", args_list[c("n1", "m1", "sd1", "n2", "m2", "sd2")]) }
      else if (method == "NNT from d") {
        cer_val <- if("cer" %in% names(row) && !is.na(row$cer)) row$cer else NULL
        res <- run_dmetar_conversion("NNT", list(d=row$d, CER=cer_val))
      }
      
      # Store results consistently
      results_list[[i]] <- res # Store the full list result
    }
    
    # Combine results with original data
    results_df <- do.call(rbind.data.frame, results_list)
    # Round numeric columns in the results part
    numeric_cols_results <- sapply(results_df, is.numeric)
    results_df[numeric_cols_results] <- lapply(results_df[numeric_cols_results], round, 4)
    cbind(df_orig, results_df) # Combine with original case data
  })
  
  
  # Display batch results table
  output$batchResultsTable <- renderDT({
    req(batchProcessedData())
    datatable(batchProcessedData(), options = list(pageLength = 10, scrollX = TRUE, autoWidth = TRUE))
  })
  
  # Download batch results
  output$downloadBatchResults <- downloadHandler(
    filename = function() { paste0("batch_converted_", gsub("[^A-Za-z0-9]", "_", input$batchMethod), "_", Sys.Date(), ".csv") },
    content = function(file) {
      req(batchProcessedData())
      write.csv(batchProcessedData(), file, row.names = FALSE)
    }
  )
  
  # --- Sample Downloads ---
  
  # Generate individual download handlers for each sample type
  lapply(allConvTypes, function(conv) {
    output[[paste0("dl_sample_", gsub("[^A-Za-z0-9]", "_", conv))]] <- downloadHandler(
      filename = function() { paste0("sample_", gsub("[^A-Za-z0-9]", "_", conv), ".csv") },
      content = function(file) {
        sample_df <- generateSampleData(conv)
        write.csv(sample_df, file, row.names = FALSE, na = "")
      }
    )
  })
  
  
  # Download all sample CSVs as ZIP
  output$downloadAllSamples <- downloadHandler(
    filename = function() { paste0("all_sample_csv_", Sys.Date(), ".zip") },
    content = function(file) {
      tmpDir <- tempdir()
      filePaths <- c()
      for(conv in allConvTypes){
        sampleData <- generateSampleData(conv)
        # Ensure all potential columns exist for consistency, even if NA
        # Create a comprehensive list of ALL possible input column names used across functions
        all_input_cols <- unique(c(
          "study", "median", "iqr", "n", "min_val", "max_val", "median_val", "q1", "q3",
          "grp1m", "grp1se", "grp1n", "grp2m", "grp2se", "grp2n", "es.type",
          "b", "sdy", "beta", "r", "f", "t",
          "effect_size", "p", "N", "effect.size.type",
          "chisq", "totaln",
          "n1", "m1", "sd1", "n2", "m2", "sd2",
          "d", "CER",
          "ai", "bi", "ci", "di", "measure"
        ))
        
        # Add missing columns with NA
        missing_cols_in_sample <- setdiff(all_input_cols, names(sampleData))
        for(col in missing_cols_in_sample) {
          sampleData[[col]] <- NA
        }
        # Keep only relevant columns for this specific sample type + study
        current_req_cols <- c("study", switch(conv,
                                              "Simple Approx (Median/IQR)" = c("median", "iqr", "n"),
                                              "Hozo et al. (2005)" = c("min_val", "max_val", "median_val", "n"),
                                              "Wan et al. (2014)" = c("q1", "median_val", "q3", "n", "min_val", "max_val"),
                                              "Mean & SE to d/g" = c("grp1m", "grp1se", "grp1n", "grp2m", "grp2se", "grp2n", "es.type"),
                                              "Unstd Reg Coef (b) to d/r" = c("b", "sdy", "grp1n", "grp2n", "es.type"),
                                              "Std Reg Coef (beta) to d/r" = c("beta", "sdy", "grp1n", "grp2n", "es.type"),
                                              "Point-biserial r to d/r" = c("r", "grp1n", "grp2n", "es.type"),
                                              "ANOVA F-value to d/g" = c("f", "grp1n", "grp2n", "es.type"),
                                              "t-Test to d/g" = c("t", "grp1n", "grp2n", "es.type"),
                                              "p-value to SE" = c("effect_size", "p", "N", "effect.size.type"),
                                              "Chi-squared to OR" = c("chisq", "totaln", "es.type"),
                                              "Pool Groups (Mean/SD)" = c("n1", "m1", "sd1", "n2", "m2", "sd2"),
                                              "NNT from d" = c("d", "CER"),
                                              "2x2 Table to OR/RR/RD" = c("ai", "bi", "ci", "di", "measure"),
                                              c()
        ))
        # Ensure 'study' is always included if available
        if("study" %in% names(sampleData)) current_req_cols <- unique(c("study", current_req_cols))
        
        sampleData <- sampleData[, intersect(names(sampleData), current_req_cols), drop=FALSE]
        
        
        csvFile <- file.path(tmpDir, paste0(gsub("[^A-Za-z0-9]", "_", conv), "_sample.csv"))
        write.csv(sampleData, csvFile, row.names = FALSE, na = "")
        filePaths <- c(filePaths, csvFile)
      }
      # Create zip file
      zip::zip(zipfile = file, files = basename(filePaths), root = tmpDir)
    },
    contentType = "application/zip"
  )
  
  # --- History ---
  output$historyTable <- renderDT({
    datatable(convHistory(), options = list(pageLength = 5, scrollX = TRUE, order = list(0, 'desc'))) # Order by time descending
  })
  
  observeEvent(input$clearHistory, {
    convHistory(data.frame(
      Time = character(),
      ConversionType = character(),
      Inputs = character(),
      Result = character(),
      stringsAsFactors = FALSE
    ))
    showNotification("History cleared", type = "message")
  })
  
}

shinyApp(ui, server)
