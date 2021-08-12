# loading packages which we use
pkg <- c("readxl", "ggplot2", "dplyr", "gridExtra", "grid")
lapply(pkg, library, character.only = TRUE)

# import data set
NTU <- read_excel("/Users/chi-yuan/Documents/NTU ECON/NTU Statistics Report/NTU_Master_Admission.xlsx",
                  sheet = "Business",
                  skip = 1)

EnrollRatioPlot <- function(file, department, title){
  data <- subset(NTU, 系所 == department)
  Interview <- data.frame(Year = data$年份,
                          Value = data$推甄錄取本校人數/data$推甄錄取人數,
                          Type = rep("Interview", nrow(data)))

  Examination <- data.frame(Year = data$年份,
                            Value = data$一般錄取本校人數/data$一般錄取人數,
                            Type = rep("Examination", nrow(data)))

  Application <- rbind(Interview, Examination)

  figure <- ggplot(data = Application,
                   aes(x = Year, y = Value, group = Type, colour = Type, na.rm = TRUE)) +
    geom_line() +
    geom_point() +
    xlab("Year") +
    ylab("Ratio") +
    xlim(2000, 2020) +
    ylim(0, 1) +
    theme_bw() +
    ggtitle(paste(title))
    #ggtitle(paste("The Ratio of Regular Students Enrolled", "in", title,
    #              "from NTU \n (Source: NTU Statistics Report)",sep = " "))

  return(figure)

}

grid.arrange(EnrollRatioPlot(NTU, "經濟學系", "Economics"),
             EnrollRatioPlot(NTU, "財務金融學系", "Finance"),
             EnrollRatioPlot(NTU, "國際企業學系", "International Business"),
             EnrollRatioPlot(NTU, "商學研究所", "Business Administration"),
             ncol = 2,
             top = textGrob("The Ratio of Regular Students Enrolled from NTU \n (Source: NTU Statistics Report)",
                            gp = gpar(fontsize=20),
                            hjust = 0.5))

######

ApplicantPlot <- function(file, department, title){
  data <- subset(NTU, 系所 == department)
  Interview <- data.frame(Year = data$年份,
                          Value = data$推甄報名人數,
                          Type = rep("Interview", nrow(data)))

  Examination <- data.frame(Year = data$年份,
                            Value = data$一般報名人數,
                            Type = rep("Examination", nrow(data)))

  Application <- rbind(Interview, Examination)

  figure <- ggplot(data = Application,
                   aes(x = Year, y = Value, group = Type, colour = Type, na.rm = TRUE)) +
    geom_line() +
    geom_point() +
    xlab("Year") +
    ylab("Number of Applicants") +
    xlim(2000, 2020) +
    ylim(0, 1500) +
    #scale_x_continuous(breaks = seq(2000, 2019, by = 5)) +
    #scale_y_continuous(breaks = seq(0, 1500, by = 100)) +
    theme_bw() +
    ggtitle(paste(title))
  #ggtitle(paste("The Ratio of Regular Students Enrolled", "in", title,
  #              "from NTU \n (Source: NTU Statistics Report)",sep = " "))

  return(figure)

}


grid.arrange(ApplicantPlot(NTU, "經濟學系", "Economics"),
             ApplicantPlot(NTU, "財務金融學系", "Finance"),
             ApplicantPlot(NTU, "國際企業學系", "International Business"),
             ApplicantPlot(NTU, "商學研究所", "Business Administration"),
             ncol = 2,
             top = textGrob("The Number of Applicants \n (Source: NTU Statistics Report)",
                            gp = gpar(fontsize=20),
                            hjust = 0.5))


