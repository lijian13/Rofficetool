
##' Write a data.frame to a excel range.
##' 
##' @title Write a data.frame to a excel range.
##' @param df The data.frame.
##' @param excelfile Path of a excel file.
##' @param shtindex Sheet index.
##' @param rangelefttop The range string, such as "A1".
##' @return TRUE or FALSE.
##' @author Jian Li <\email{rweibo@@sina.com}>

writeDfToRange <- function(dF, excelfile, shtindex, rangelefttop) {
    xlsfile <- normalizePath(excelfile, winslash = "/", mustWork = TRUE)
    excelapp <- COMCreate("Excel.Application")
    excelwbk <- excelapp[["workbooks"]]$Open(xlsfile)
    
    sht1 <- excelwbk[["sheets"]]$Item(shtindex)
    for (i in 1:nrow(dF)) {
        for (j in 1:ncol(dF)) {
            cell.tmp <- sht1[["Cells"]]$Item(rangelefttop[1] + i - 1, rangelefttop[2] + j -1)
            if (!is.na(dF[i, j])) {
                cell.tmp[["Value"]] <- dF[i, j]
            }
        }
        
    }
    excelwbk$Close(TRUE)
}

