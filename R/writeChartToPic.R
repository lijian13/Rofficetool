
##' Write a Excel chart to a picture file.
##' 
##' @title Write a Excel chart to a picture file.
##' @param excelfile Path of a excel file.
##' @param shtindex Sheet index.
##' @param chartindex Chart index.
##' @param picfile Path of the picture file.
##' @return TRUE or FALSE
##' @author Jian Li <\email{rweibo@@sina.com}>

writeChartToPic <- function(excelfile, shtindex, chartindex, picfile) {
    xlsfile <- normalizePath(excelfile, winslash = "/", mustWork = TRUE)
	picfile <- normalizePath(picfile, winslash = "/", mustWork = FALSE)
    excelapp <- COMCreate("Excel.Application")
    excelwbk <- excelapp[["workbooks"]]$Open(xlsfile)
    
    sht1 <- excelwbk[["sheets"]]$Item(shtindex)
	sht1$ChartObjects()$Item(chartindex)[["Chart"]]$Export(picfile)
    excelwbk$Close(FALSE)
	return(picfile)
}

