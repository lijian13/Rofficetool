

require(Rofficetool)


setwd("E:\\Mango\\Training\\Youku\\20140530\\examples\\vba")

library(Rofficetool)

templateppt <- system.file("examples", "pptdemo.ppt", package = "Rofficetool")
templatexls <- system.file("examples", "exceldemo.xls", package = "Rofficetool")

thisppt <- paste0("ppt", "_", format(Sys.time(), "%Y%m%d%H%M%S"), ".ppt")
thisxls <- paste0("xls", "_", format(Sys.time(), "%Y%m%d%H%M%S"), ".xls")

# copy template
thisppt <- copyFromTemplate(templateppt, thisppt)
thisxls <- copyFromTemplate(templatexls, thisxls)

# parse ppt
p1 <- parsePPT(thisppt)

# write data.frame to range of excel
d1 <- iris
d2 <- d1[sample(1:nrow(d1), nrow(d1)), ]
d2$Species <- as.character(d2$Species)

writeDfToRange(d2, thisxls, shtindex = 1, c(2,1))


# write range to pic

pic1 <- writeRangeToPic(thisxls, shtindex = 2, rangestring = "B2:G10", 
	picfile = "p1.jpg", chartsize = c(300, 150))

# write chart to pic

pic2 <- writeChartToPic(thisxls, shtindex = 3, chartindex = 1, 
	picfile = "p2.jpg")

# write pic to ppt

writePicToPPT(thisppt, picfile = pic1, ipage = 4, 
	Position = c(100, 200), Size = c(540, 280))

writePicToPPT(thisppt, picfile = pic2, ipage = 3, 
	Position = c(100, 200), Size = c(540, 280))


# write data.frame to ppt
d3 <- d2[1:8, ]
d3 <- cbind(rep(NA, nrow(d3)), d3)
d3 <- rbind(rep(NA, ncol(d3)), d3)

writeDfToPPT(d3, thisppt, ipage = 5, ishape = 2) 


