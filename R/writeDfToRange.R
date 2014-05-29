
##' Write a data.frame to a excel range.
##' 
##' @title Write a data.frame to a excel range.
##' @param df the data.frame
##' @param excelfile path of a excel file
##' @param shtindex sheet index
##' @param rangelefttop the range string, such as "A1"
##' @return TRUE or FALSE
##' @author Jian Li <\email{rweibo@@sina.com}>


#dF <- tdf2
#excelfile <- "E:/Mango/Training/Youku/20140515/examples/report/table_1.xlsx"
#shtindex <- 1
#rangelefttop <- c(5, 2)
# writeDfToRange( tdf2, "E:/Mango/Training/Youku/20140515/examples/report/table_1.xlsx", 1, c(5,2))

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

