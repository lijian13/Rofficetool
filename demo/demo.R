

library(Rofficetool)

templateppt <- system.file("examples", "pptdemo.ppt", package = "Rofficetool")
templatexls <- system.file("examples", "exceldemo.xls", package = "Rofficetool")

# copy template
thisppt <- setOutput(templateppt)
thisxls <- setOutput(templatexls)

# read ppt file
ppt1 <- readPPT(thisppt)
class(ppt1)

# parse ppt
p1 <- parsePPT(ppt1)
p1

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

writePicToPPT(picfile = pic1, ppt1, ipage = 4, horizontal = 3.5, vertical = 7, height = 10, width = 20)

writePicToPPT(picfile = pic2, ppt1, ipage = 3, horizontal = 3.5, vertical = 7, height = 10, width = 20)



# write data.frame to ppt
d3 <- matrix(round(rnorm(40), 2), 8, 5)
writeDfToPPT(Df = d3, ppt1, ipage = 5, ishape = 2, startrow = 2, startcol = 2)


# write text to ppt
writeTextToPPT(as.character(Sys.Date()), ppt1, ipage = 1, ishape = 2)


