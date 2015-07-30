
##' Write excel range to a picture file.
##' 
##' @title Write excel range to a picture file.
##' @param excelfile Path of a excel file.
##' @param shtindex Sheet index.
##' @param rangestring The range string, such as "A1".
##' @param picfile The path of the picture.
##' @param chartsize Chart size.
##' @return TRUE or FALSE.
##' @author Jian Li <\email{rweibo@@sina.com}>

writeRangeToPic <- function(excelfile, shtindex, rangestring, picfile, chartsize = c(800, 600)) {
    xlsfile <- normalizePath(excelfile, winslash = "/", mustWork = TRUE)
	picfile <- normalizePath(picfile, winslash = "/", mustWork = FALSE)
    excelapp <- COMCreate("Excel.Application")
    excelwbk <- excelapp[["workbooks"]]$Open(xlsfile)
    
    sht1 <- excelwbk[["sheets"]]$Item(shtindex)
    sht1$Range(rangestring)$CopyPicture()
	sht1$ChartObjects()$Add(0, 0, chartsize[1], chartsize[2])
	nchart <- sht1$ChartObjects()$Count()
	sht1$ChartObjects()$Item(nchart)[["Chart"]]$Paste()
	sht1$ChartObjects()$Item(nchart)[["Chart"]]$Export(picfile)
    excelwbk$Close(FALSE)
	return(picfile)
}

