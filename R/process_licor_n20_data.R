#' Process LICOR N2O data with interactive BI/AI selection
#'
#' This function processes LICOR greenhouse gas analyzer data and allows the
#' user to interactively select regions before injection (BI) and after
#' injection (AI) within specified time windows. The selected regions are used
#' to calculate average N2O concentrations and diagnostic values. Plots of the
#' selected regions are saved and results are exported to an Excel file.
#'
#' @param data_folder Character. Path to the folder containing LICOR `.data`
#'   files.
#' @param plot_folder Character. Path to the folder where output plots will
#'   be saved.
#' @param tw_csv Character. Path to a CSV file containing time windows with
#'   columns `site`, `start`, and `end`.
#' @param window_indices Numeric vector specifying which time windows to
#'   process (e.g., `1:5`). If `NULL`, all windows are processed.
#' @param excel_file_base Character. Base name for the Excel file where results
#'   will be saved.
#'
#' @details
#' The function performs the following steps:
#' \enumerate{
#'   \item Reads LICOR `.data` files from the specified folder.
#'   \item Loads time windows from the provided CSV file.
#'   \item Plots N2O concentration against time for each selected window.
#'   \item Prompts the user to select the BI (before injection) region.
#'   \item Prompts the user to select the AI (after injection) region.
#'   \item Calculates mean N2O concentrations and mean DIAG values.
#'   \item Saves annotated plots showing the selected regions.
#'   \item Exports results to an Excel file.
#' }
#'
#' Selected BI and AI regions are highlighted on the plot using shaded
#' rectangles for visual confirmation.
#'
#' @return A data frame containing:
#' \describe{
#'   \item{Site}{Site name}
#'   \item{Date}{Measurement date}
#'   \item{bi_start}{Start time of BI region}
#'   \item{bi_end}{End time of BI region}
#'   \item{ai_start}{Start time of AI region}
#'   \item{ai_end}{End time of AI region}
#'   \item{Avg_bi}{Average N2O concentration before injection}
#'   \item{Avg_ai}{Average N2O concentration after injection}
#'   \item{Avg_DIAG_bi}{Average DIAG value for BI region}
#'   \item{Avg_DIAG_ai}{Average DIAG value for AI region}
#' }
#'
#' @author Oswald Omuron
#'
#' @importFrom graphics locator par points rect
#' @importFrom grDevices png dev.off adjustcolor
#' @importFrom utils read.csv read.table
#'
#' @export


process_licor_n20_data <- function(
    data_folder,
    plot_folder,
    tw_csv,
    window_indices = NULL,
    excel_file_base = "LICOR_N2O_results.xlsx"
) {
  if(!dir.exists(plot_folder)) dir.create(plot_folder, recursive = TRUE)

  # Load time windows from CSV
  tw_df <- read.csv(tw_csv, stringsAsFactors = FALSE, skip = 1)
  tw_df$start <- as.POSIXct(tw_df$start, format="%Y/%m/%d %H:%M:%S", tz="UTC")
  tw_df$end   <- as.POSIXct(tw_df$end,   format="%Y/%m/%d %H:%M:%S", tz="UTC")

  time_windows <- lapply(1:nrow(tw_df), function(i) {
    list(
      site  = tw_df$site[i],
      start = tw_df$start[i],
      end   = tw_df$end[i]
    )
  })

  if(is.null(window_indices)) window_indices <- seq_along(time_windows)

  # -----------------------------
  # Function to read and clean LICOR file (keep DIAG)
  read_clean_file <- function(file_path){
    d <- read.table(file_path, sep="\t", skip=5, header=TRUE)
    d <- d[,c("DATE","TIME","N2O","DIAG")]
    d <- d[-1,]
    d$date_time <- paste(d$DATE,d$TIME)
    d$date <- as.POSIXct(d$date_time, format="%Y-%m-%d %H:%M:%OS", tz="UTC")
    d <- d[!duplicated(d$date),]
    d$N2O_ppb <- as.numeric(d$N2O)
    d$DIAG <- as.numeric(d$DIAG)
    return(d[,c("date","N2O_ppb","DIAG")])
  }

  # Load all LICOR files
  files <- list.files(data_folder, pattern="\\.data$", full.names=TRUE)
  all_data <- do.call(rbind, lapply(files, read_clean_file))

  # -----------------------------
  # Function to select region interactively
  select_region <- function(plot_data, col="red"){

    cat("Click START of region (or Esc to skip)\n")
    p1 <- locator(1)
    if(is.null(p1)) return(NULL)

    cat("Click END of region\n")
    p2 <- locator(1)
    if(is.null(p2)) return(NULL)

    # Convert locator x values back to POSIXct
    start_time <- as.POSIXct(p1$x, origin="1970-01-01", tz="UTC")
    end_time   <- as.POSIXct(p2$x, origin="1970-01-01", tz="UTC")

    usr <- par("usr")

    rect(min(start_time,end_time), usr[3],
         max(start_time,end_time), usr[4],
         col=adjustcolor(col, alpha.f=0.25), border=col)


    # Pause so user can confirm selection
    Sys.sleep(3)

    selected <- plot_data[
      plot_data$date >= min(start_time,end_time) &
        plot_data$date <= max(start_time,end_time), ]

    return(selected)
  }

  # -----------------------------
  # Initialize results
  results <- data.frame()
  processed_windows <- c()  # track processed windows

  # -----------------------------
  # Loop through selected windows
  for(i in window_indices){
    site_name <- time_windows[[i]]$site
    start_time <- as.POSIXct(time_windows[[i]]$start, tz="UTC")
    end_time   <- as.POSIXct(time_windows[[i]]$end, tz="UTC")

    plot_data <- all_data[all_data$date >= start_time & all_data$date <= end_time, ]
    if(nrow(plot_data)==0) next

    plot(plot_data$date, plot_data$N2O_ppb,
         type="l",
         col="blue",
         xlab="Time",
         ylab="N2O (ppb)",
         xaxt="n",
         main=paste(site_name,
                    "- Select the before injection region, then the after injection region"))


    axis.POSIXct(1, at=pretty(plot_data$date), format="%H:%M:%S")


    # BI region
    cat("Select the before injection (bi) region\n")
    bi_data <- select_region(plot_data, col="red")
    if(!is.null(bi_data)){
      Avg_bi <- mean(bi_data$N2O_ppb, na.rm=TRUE)
      Avg_DIAG_bi <- mean(bi_data$DIAG, na.rm=TRUE)
      bi_start <- format(min(bi_data$date), "%H:%M:%S")
      bi_end   <- format(max(bi_data$date), "%H:%M:%S")
    } else {
      Avg_bi <- Avg_DIAG_bi <- bi_start <- bi_end <- NA
    }

    # AI region
    cat("Select the after injection (ai) region\n")
    ai_data <- select_region(plot_data, col="darkgreen")
    if(!is.null(ai_data)){
      Avg_ai <- mean(ai_data$N2O_ppb, na.rm=TRUE)
      Avg_DIAG_ai <- mean(ai_data$DIAG, na.rm=TRUE)
      ai_start <- format(min(ai_data$date), "%H:%M:%S")
      ai_end   <- format(max(ai_data$date), "%H:%M:%S")
    } else {
      Avg_ai <- Avg_DIAG_ai <- ai_start <- ai_end <- NA
    }

    # Save plot
    plot_file <- paste0(plot_folder, "/", site_name, "_window_", i, ".png")
    png(plot_file, width=1200, height=800, res=150)
    plot(plot_data$date, plot_data$N2O_ppb, type="l", col="blue",
         xlab="Time", ylab="N2O (ppb)",
         xaxt="n",
         main=paste(site_name,
                    "\nbefore injection =", round(Avg_bi,2), "ppb",
                    "\nafter injection =", round(Avg_ai,2), "ppb"))

    axis.POSIXct(1, at=pretty(plot_data$date), format="%H:%M:%S",mgp = c(0,0.5,0), tck = -0.02)


    usr <- par("usr")
    if(!is.null(bi_data)) rect(min(bi_data$date), usr[3], max(bi_data$date), usr[4],
                               col=adjustcolor("red", alpha.f=0.25), border="red")
    if(!is.null(ai_data)) rect(min(ai_data$date), usr[3], max(ai_data$date), usr[4],
                               col=adjustcolor("darkgreen", alpha.f=0.25), border="darkgreen")
    if(!is.null(bi_data)) points(bi_data$date, bi_data$N2O_ppb, col="red", pch=20)
    if(!is.null(ai_data)) points(ai_data$date, ai_data$N2O_ppb, col="darkgreen", pch=20)
    dev.off()

    # Save results
    row_data <- data.frame(
      Site = site_name,
      Date = as.Date(plot_data$date[1]),
      bi_start = bi_start,
      bi_end   = bi_end,
      ai_start = ai_start,
      ai_end   = ai_end,
      Avg_bi = Avg_bi,
      Avg_ai = Avg_ai,
      Avg_DIAG_bi = Avg_DIAG_bi,
      Avg_DIAG_ai = Avg_DIAG_ai
    )
    results <- rbind(results, row_data)

    # Track and print processed window in console
    processed_windows <- c(processed_windows, i)
    cat("✅ Processed window:", i, "Site:", site_name, "\n")
  }

  # -----------------------------
  # Export results to Excel with window numbers
  if(!require(openxlsx, quietly=TRUE)) install.packages("openxlsx")
  library(openxlsx)

  window_str <- paste(processed_windows, collapse = "-")
  excel_file <- sub(".xlsx$", paste0("_windows_", window_str, ".xlsx"), excel_file_base)

  write.xlsx(results, file=excel_file, overwrite = TRUE)
  cat("\nProcessing complete. Results saved to:", excel_file, "\n")
  cat("All processed windows:", paste(processed_windows, collapse = ", "), "\n")


  return(results)
}
