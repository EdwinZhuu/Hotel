# HOTEL project
library(readxl)
library(xlsx)
library(dplyr)
library(leaps)

#setwd("C:/Users/edwin/Desktop/Volunteer/Datasets/")
setwd("C:/Users/Edwin/Desktop/Volunteer/Hotel/")

subjects = read_excel("Baseline cSVD Demo.xlsx", sheet = "demo")
binary = read_excel("Baseline cSVD Demo.xlsx", sheet = "binary", skip = 1)
weighted = read_excel("Baseline cSVD Demo.xlsx", sheet = "weighted", skip = 1)
colnames(binary)[1] = "Subject"
colnames(weighted)[1] = "Subject"

binary = merge(subjects, binary, by = "Subject")
binary = subset(binary, Subject != "01_083A" & Subject != "01_085A" & Subject != "01_126A" 
                & Subject != "02_012A" & Subject != "04_013A" & Subject != "03_029A")
colnames(binary) = sub("\\.\\.\\..*", "", colnames(binary))
binary_split = split(binary, binary$Group)

weighted = merge(subjects, weighted, by = "Subject")
weighted = subset(weighted, Subject != "01_083A" & Subject != "01_085A" & Subject != "01_126A" 
                & Subject != "02_012A" & Subject != "04_013A" & Subject != "03_029A")
colnames(weighted) = sub("\\.\\.\\..*", "", colnames(weighted))
weighted_split = split(weighted, weighted$Group)

binary_control = binary_split$Control
binary_SVD = binary_split$SVD

weighted_control = weighted_split$Control
weighted_SVD = weighted_split$SVD

binary_control_matrix = matrix(ncol = 14, nrow = 17)
colnames(binary_control_matrix) = colnames(binary_control[6:19])
rownames(binary_control_matrix) = seq(from=0.08, to=0.40, by=0.02)

binary_SVD_matrix = matrix(ncol = 14, nrow = 17)
colnames(binary_SVD_matrix) = colnames(binary_SVD[6:19])
rownames(binary_SVD_matrix) = seq(from=0.08, to=0.40, by=0.02)

# Check for normal distribution in binary
for (i in 1:17){
  x = 6 + (i-1)*15
  for (j in 1:14){
    test_control = shapiro.test(binary_control[,x])
    test_SVD = shapiro.test(binary_SVD[,x])
    binary_control_matrix[i,j] = test_control$p.value
    binary_SVD_matrix[i,j] = test_SVD$p.value
    x = x+1
  }
}

binary_control_matrix
binary_SVD_matrix

empty_row = matrix(ncol = 14, nrow = 1)
all_shapiro = rbind(binary_control_matrix, empty_row, empty_row)
all_shapiro = rbind(all_shapiro, binary_SVD_matrix)

write.xlsx(all_shapiro, "Output.xlsx", sheetName = "Binary Shapiro-Wilk",
                      row.names = TRUE, col.names = TRUE)

# Test the AIC and R squared
binary_control_matrix_intercept = matrix(ncol = 14, nrow = 17)
colnames(binary_control_matrix_intercept) = colnames(binary_control[6:19])
rownames(binary_control_matrix_intercept) = seq(from=0.08, to=0.40, by=0.02)

binary_control_matrix_age = matrix(ncol = 14, nrow = 17)
colnames(binary_control_matrix_age) = colnames(binary_control[6:19])
rownames(binary_control_matrix_age) = seq(from=0.08, to=0.40, by=0.02)

binary_control_matrix_gender = matrix(ncol = 14, nrow = 17)
colnames(binary_control_matrix_gender) = colnames(binary_control[6:19])
rownames(binary_control_matrix_gender) = seq(from=0.08, to=0.40, by=0.02)

binary_control_matrix_AG = matrix(ncol = 14, nrow = 17)
colnames(binary_control_matrix_AG) = colnames(binary_control[6:19])
rownames(binary_control_matrix_AG) = seq(from=0.08, to=0.40, by=0.02)

binary_control_matrix_education = matrix(ncol = 14, nrow = 17)
colnames(binary_control_matrix_education) = colnames(binary_control[6:19])
rownames(binary_control_matrix_education) = seq(from=0.08, to=0.40, by=0.02)

for (a in 1:17){
  b = 6 + (a-1)*15
  for (c in 1:14){
    control_fit_1 = lm(binary_control[,b] ~ 1)
    control_fit_2 = lm(binary_control[,b] ~ binary_control$Age)
    control_fit_3 = lm(binary_control[,b] ~ binary_control$Gender)
    control_fit_4 = lm(binary_control[,b] ~ binary_control$Age + binary_control$Gender)
    control_fit_5 = lm(binary_control[,b] ~ binary_control$Age + binary_control$Gender + binary_control$Education)
    binary_control_matrix_intercept[a,c] = summary(control_fit_1)$adj.r.squared
    binary_control_matrix_age[a,c] = summary(control_fit_2)$adj.r.squared
    binary_control_matrix_gender[a,c] = summary(control_fit_3)$adj.r.squared
    binary_control_matrix_AG[a,c] = summary(control_fit_4)$adj.r.squared
    binary_control_matrix_education[a,c] = summary(control_fit_5)$adj.r.squared
    b = b+1
  }
}

empty_col = matrix(nrow = 17, ncol = 1)
all_normal_lm = cbind(binary_control_matrix_intercept, empty_col, binary_control_matrix_age,
                      empty_col, binary_control_matrix_gender, empty_col, binary_control_matrix_AG,
                      empty_col, binary_control_matrix_education)

write.xlsx(all_normal_lm, "Output.xlsx", sheetName = "Linear model for Binary Control",
           row.names = TRUE, col.names = TRUE, append = TRUE)

binary_SVD_matrix_intercept = matrix(ncol = 14, nrow = 17)
colnames(binary_SVD_matrix_intercept) = colnames(binary_SVD[6:19])
rownames(binary_SVD_matrix_intercept) = seq(from=0.08, to=0.40, by=0.02)

binary_SVD_matrix_age = matrix(ncol = 14, nrow = 17)
colnames(binary_SVD_matrix_age) = colnames(binary_SVD[6:19])
rownames(binary_SVD_matrix_age) = seq(from=0.08, to=0.40, by=0.02)

binary_SVD_matrix_gender = matrix(ncol = 14, nrow = 17)
colnames(binary_SVD_matrix_gender) = colnames(binary_SVD[6:19])
rownames(binary_SVD_matrix_gender) = seq(from=0.08, to=0.40, by=0.02)

binary_SVD_matrix_AG = matrix(ncol = 14, nrow = 17)
colnames(binary_SVD_matrix_AG) = colnames(binary_SVD[6:19])
rownames(binary_SVD_matrix_AG) = seq(from=0.08, to=0.40, by=0.02)

binary_SVD_matrix_education = matrix(ncol = 14, nrow = 17)
colnames(binary_SVD_matrix_education) = colnames(binary_SVD[6:19])
rownames(binary_SVD_matrix_education) = seq(from=0.08, to=0.40, by=0.02)

for (a in 1:17){
  b = 6 + (a-1)*15
  for (c in 1:14){
    SVD_fit_1 = lm(binary_SVD[,b] ~ 1)
    SVD_fit_2 = lm(binary_SVD[,b] ~ binary_SVD$Age)
    SVD_fit_3 = lm(binary_SVD[,b] ~ binary_SVD$Gender)
    SVD_fit_4 = lm(binary_SVD[,b] ~ binary_SVD$Age + binary_SVD$Gender)
    SVD_fit_5 = lm(binary_SVD[,b] ~ binary_SVD$Age + binary_SVD$Gender + binary_SVD$Education)
    binary_SVD_matrix_intercept[a,c] = summary(SVD_fit_1)$adj.r.squared
    binary_SVD_matrix_age[a,c] = summary(SVD_fit_2)$adj.r.squared
    binary_SVD_matrix_gender[a,c] = summary(SVD_fit_3)$adj.r.squared
    binary_SVD_matrix_AG[a,c] = summary(SVD_fit_4)$adj.r.squared
    binary_SVD_matrix_education[a,c] = summary(SVD_fit_5)$adj.r.squared
    b = b+1
  }
}

binary_SVD_matrix_intercept
binary_SVD_matrix_age
binary_SVD_matrix_gender
binary_SVD_matrix_AG
binary_SVD_matrix_education

all_normal_SVD_lm = cbind(binary_SVD_matrix_intercept, empty_col, binary_SVD_matrix_age,
                      empty_col, binary_SVD_matrix_gender, empty_col, binary_SVD_matrix_AG,
                      empty_col, binary_SVD_matrix_education)

write.xlsx(all_normal_SVD_lm, "Output.xlsx", sheetName = "Linear model for Binary SVD",
           row.names = TRUE, col.names = TRUE, append = TRUE)

# Weighted distribution
weighted_control_matrix = matrix(ncol = 14, nrow = 1)
colnames(weighted_control_matrix) = colnames(weighted_control[6:19])
rownames(weighted_control_matrix) = 0.20

weighted_SVD_matrix = matrix(ncol = 14, nrow = 1)
colnames(weighted_SVD_matrix) = colnames(weighted_SVD[6:19])
rownames(weighted_SVD_matrix) = 0.20

# Check for normal distribution in weighted
for (j in 1:14){
  x = 6
  weighted_test_control = shapiro.test(weighted_control[,x])
  weighted_test_SVD = shapiro.test(weighted_SVD[,x])
  weighted_control_matrix[1,j] = weighted_test_control$p.value
  weighted_SVD_matrix[1,j] = weighted_test_SVD$p.value
  x = x+1
}

weighted_control_matrix
weighted_SVD_matrix

all_weighted_distributions = rbind(weighted_control_matrix, empty_row, weighted_SVD_matrix)
write.xlsx(all_weighted_distributions, "Output.xlsx", sheetName = "Weighted Shapiro Wilk Test",
           row.names = TRUE, col.names = TRUE, append = TRUE)


weighted_control_matrix_lm = matrix(ncol = 14, nrow = 5)
colnames(weighted_control_matrix_lm) = colnames(weighted_SVD[6:19])
rownames(weighted_control_matrix_lm) = c("Intercept", "Age", "Gender", "AG", "Education")

weighted_SVD_matrix_lm = matrix(ncol = 14, nrow = 5)
colnames(weighted_SVD_matrix_lm) = colnames(weighted_SVD[6:19])
rownames(weighted_SVD_matrix_lm) = c("Intercept", "Age", "Gender", "AG", "Education")

for (j in 1:14){
  x = 6
  control_test1 = lm(weighted_control[,x] ~ 1)
  control_test2 = lm(weighted_control[,x] ~ weighted_control$Age)
  control_test3 = lm(weighted_control[,x] ~ weighted_control$Gender)
  control_test4 = lm(weighted_control[,x] ~ weighted_control$Age + weighted_control$Gender)
  control_test5 = lm(weighted_control[,x] ~ weighted_control$Age + weighted_control$Gender + weighted_control$Education)
  SVD_test1 = lm(weighted_SVD[,x] ~ 1)
  SVD_test2 = lm(weighted_SVD[,x] ~ weighted_SVD$Age)
  SVD_test3 = lm(weighted_SVD[,x] ~ weighted_SVD$Gender)
  SVD_test4 = lm(weighted_SVD[,x] ~ weighted_SVD$Age + weighted_SVD$Gender)
  SVD_test5 = lm(weighted_SVD[,x] ~ weighted_SVD$Age + weighted_SVD$Gender + weighted_SVD$Education)
  weighted_control_matrix_lm[1,j] = summary(control_test1)$adj.r.squared
  weighted_control_matrix_lm[2,j] = summary(control_test2)$adj.r.squared
  weighted_control_matrix_lm[3,j] = summary(control_test3)$adj.r.squared
  weighted_control_matrix_lm[4,j] = summary(control_test4)$adj.r.squared
  weighted_control_matrix_lm[5,j] = summary(control_test5)$adj.r.squared
  weighted_SVD_matrix_lm[1,j] = summary(SVD_test1)$adj.r.squared
  weighted_SVD_matrix_lm[2,j] = summary(SVD_test2)$adj.r.squared
  weighted_SVD_matrix_lm[3,j] = summary(SVD_test3)$adj.r.squared
  weighted_SVD_matrix_lm[4,j] = summary(SVD_test4)$adj.r.squared
  weighted_SVD_matrix_lm[5,j] = summary(SVD_test5)$adj.r.squared
  x = x+1
}

weighted_control_matrix_lm
weighted_SVD_matrix_lm

all_weighted_lm = rbind(weighted_control_matrix_lm, empty_row, weighted_SVD_matrix_lm)

write.xlsx(all_weighted_lm, "Output.xlsx", sheetName = "Weighted linear models",
           row.names = TRUE, col.names = TRUE, append = TRUE)

# Hybrid Stepwise Function
fit.start = lm(binary_control[,9] ~ 1)
fit.end = lm(binary_control[,9] ~ binary_control$Age + binary_control$Gender + binary_control$Education)

step.AIC = step(fit.start, list(upper = fit.end), k = 2)

extractAIC(step.AIC)[2]

all.densities = replicate(17,data.frame(Parameter = character(0),
                  AIC = double(0),
                  `Linear Model` = character(0),
                  stringsAsFactors = FALSE), simplify = FALSE)
names(all.densities) <- paste0('density', seq(from = 0.08, to = 0.40, by=0.02))
all.densities

for (i in 1:17){
  x = 6 + (i-1)*15
  for (j in 1:14){
    fit.start = lm(binary_control[,x] ~ 1)
    fit.end = lm(binary_control[,x] ~ binary_control$Age + binary_control$Gender + binary_control$Education)
    
    step.AIC = step(fit.start, list(upper = fit.end), k = 2)
    names = paste(colnames(model.matrix(step.AIC)),collapse=" ")
    names = gsub("binary_control\\$", "", names, perl = TRUE)
    names = gsub("\\(", "", names, perl = TRUE)
    names = gsub("\\)", "", names, perl = TRUE)
    
    all.densities[[i]][j,1] = colnames(binary_control)[x]
    all.densities[[i]][j,2] = extractAIC(step.AIC)[2]
    all.densities[[i]][j,3] = names
    
    x = x+1
  }
}

all.densities.combined = do.call("cbind", all.densities)

write.xlsx(all.densities.combined, "../Hotel/Output.xlsx", sheetName = "AIC", append =TRUE) 
