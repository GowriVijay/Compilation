# install.packages("foreign")
# install.packages("sqldf")
library(foreign)
library(sqldf)
library(tcltk)

#  Use the Keyboard shortcut Ctrl+shift+H for changing the working directory

setwd("C:\\Users\\vthangamuthu\\Documents\\Vijay\\Sam\\Authentication Survey\\Authentication Survey - V2")


rm(list = ls())
# load("Attitudinal_Segment_Workspace.RData")
All <- read.spss("Authentication Survey - V2 - Aug 11th 9 PM.sav",to.data.frame = T)
# US <- subset(All,All$vco == c("us"))
UK <- subset(All,All$vco == c("uk"))
# DE <- subset(All,All$vco == c("de"))

UK_1 <-subset(UK,UK$Q2_Order == c("Current / Virtual / Physical"))
Bk_up_1 <- UK_1

UK_1_VA_Outlier_removed <- UK_1[ which((UK_1$Q5A_Q5B < 3 & UK_1$Q5A_Q5B >= 0) & (UK_1$Q2A == c("Extremely unlikely 1") | UK_1$Q2A == c("2") | UK_1$Q2A == c("3")) & (UK_1$Q2B == c("Extremely likely 5") | UK_1$Q2B == c("4"))), ] #I've added the outlier piece for UK_1$Q5A_Q5B
UK_1_PA_Outlier_removed <- UK_1[ which((UK_1$Q5A_Q5C < 3 & UK_1$Q5A_Q5C >= 0) & (UK_1$Q2A == c("Extremely unlikely 1") | UK_1$Q2A == c("2") | UK_1$Q2A == c("3")) & (UK_1$Q2C == c("Extremely likely 5") | UK_1$Q2C == c("4"))), ] #I've added the outlier piece for UK_1$Q5A_Q5C

# & (UK_1$Q2A == c("Extremely likely 5") | UK_1$Q2A == c("4"))
# & (UK_1$Q2A == c("Extremely likely 5") | UK_1$Q2A == c("4"))
#Doing the next line just because the dataframe was named as UK_1 and I'm too lazy to rename them
UK_1_VA <- UK_1_VA_Outlier_removed
UK_1_PA <- UK_1_PA_Outlier_removed


# Calculating Avg % increase for VA
A1 <- aggregate(UK_1_VA$Q5A_Q5B, by = list(Gender = UK_1_VA$Gender), FUN = length)
B1 <- aggregate(UK_1_VA$Q5A_Q5B, by = list(Vertical_Gender = UK_1_VA$Vrtcl_Gndr,Item_Condition = UK_1_VA$vchrt), FUN = length)
C1 <- aggregate(UK_1_VA$Q5A_Q5B, by = list(Vertical_Gender = UK_1_VA$Vrtcl_Gndr), FUN = length)
D1 <- aggregate(UK_1_VA$Q5A_Q5B, by = list(Gender = UK_1_VA$Gender,Item_Condition = UK_1_VA$vchrt), FUN = length)

# Calculating Avg % increase for PA
W1 <- aggregate(UK_1_PA$Q5A_Q5C, by = list(Gender = UK_1_PA$Gender), FUN = length)
X1 <- aggregate(UK_1_PA$Q5A_Q5C, by = list(Vertical_Gender = UK_1_PA$Vrtcl_Gndr,Item_Condition = UK_1_PA$vchrt), FUN = length)
Y1 <- aggregate(UK_1_PA$Q5A_Q5C, by = list(Vertical_Gender = UK_1_PA$Vrtcl_Gndr), FUN = length)
Z1 <- aggregate(UK_1_PA$Q5A_Q5C, by = list(Gender = UK_1_PA$Gender,Item_Condition = UK_1_PA$vchrt), FUN = length)



UK_2 <-subset(UK,UK$Q2_Order == c("Current / Physical / Virtual"))
Bk_up_2 <- UK_2

UK_2_VA_Outlier_removed <- UK_2[ which((UK_2$Q5A_Q5B < 3 & UK_2$Q5A_Q5B >= 0) & (UK_2$Q2A == c("Extremely unlikely 1") | UK_2$Q2A == c("2") | UK_2$Q2A == c("3")) & (UK_2$Q2B == c("Extremely likely 5") | UK_2$Q2B == c("4"))), ] #I've added the outlier piece for UK_2$Q5A_Q5B
UK_2_PA_Outlier_removed <- UK_2[ which((UK_2$Q5A_Q5C < 3 & UK_2$Q5A_Q5C >= 0) & (UK_2$Q2A == c("Extremely unlikely 1") | UK_2$Q2A == c("2") | UK_2$Q2A == c("3")) & (UK_2$Q2C == c("Extremely likely 5") | UK_2$Q2C == c("4"))), ] #I've added the outlier piece for UK_2$Q5A_Q5C

# & (UK_2$Q2A == c("Extremely likely 5") | UK_2$Q2A == c("4"))
# & (UK_2$Q2A == c("Extremely likely 5") | UK_2$Q2A == c("4"))
#Doing the next line just because the dataframe was named as UK_2 and I'm too lazy to rename them
UK_2_VA <- UK_2_VA_Outlier_removed
UK_2_PA <- UK_2_PA_Outlier_removed

# Calculating Avg % increase for PA
W2 <- aggregate(UK_2_PA$Q5A_Q5C, by = list(Gender = UK_2_PA$Gender), FUN = length)
X2 <- aggregate(UK_2_PA$Q5A_Q5C, by = list(Vertical_Gender = UK_2_PA$Vrtcl_Gndr,Item_Condition = UK_2_PA$vchrt), FUN = length)
Y2 <- aggregate(UK_2_PA$Q5A_Q5C, by = list(Vertical_Gender = UK_2_PA$Vrtcl_Gndr), FUN = length)
Z2 <- aggregate(UK_2_PA$Q5A_Q5C, by = list(Gender = UK_2_PA$Gender,Item_Condition = UK_2_PA$vchrt), FUN = length)

# Calculating Avg % increase for VA
A2 <- aggregate(UK_2_VA$Q5A_Q5B, by = list(Gender = UK_2_VA$Gender), FUN = length)
B2 <- aggregate(UK_2_VA$Q5A_Q5B, by = list(Vertical_Gender = UK_2_VA$Vrtcl_Gndr,Item_Condition = UK_2_VA$vchrt), FUN = length)
C2 <- aggregate(UK_2_VA$Q5A_Q5B, by = list(Vertical_Gender = UK_2_VA$Vrtcl_Gndr), FUN = length)
D2 <- aggregate(UK_2_VA$Q5A_Q5B, by = list(Gender = UK_2_VA$Gender,Item_Condition = UK_2_VA$vchrt), FUN = length)

# Doing the next set of work for replacing the Segments with missing information
bs <- data.frame(Vertical_Gender = rep(sort(unique(All$Vrtcl_Gndr)),3),Item_Condition = rep(sort(unique(All$vchrt)),each = 10))
cs <- data.frame(Vertical_Gender = sort(unique(trimws(as.character(All$Vrtcl_Gndr),"right"))))
bs$New <- as.factor(paste(trimws(as.character(bs$Vertical_Gender),"right"),bs$Item_Condition))


# Replacing missing infor for Accessories - Handbags - Male for VA1
B1$New <- as.factor(paste(trimws(as.character(B1$Vertical_Gender),"right"),B1$Item_Condition))
C1$Vertical_Gender <- trimws(as.character(C1$Vertical_Gender),"right")

B1 <- sqldf("SELECT bs.Vertical_Gender, bs.Item_Condition, B1.x
            FROM bs
            LEFT JOIN B1
            ON bs.New = B1.New")
B1$x[is.na(B1$x)] <- 0

C1 <- sqldf("SELECT cs.Vertical_Gender, C1.x
            FROM cs
            LEFT JOIN C1
            ON cs.Vertical_Gender = C1.Vertical_Gender")
C1$x[is.na(C1$x)] <- 0


# Replacing missing infor for Accessories - Handbags - Male for PA1
X1$New <- as.factor(paste(trimws(as.character(X1$Vertical_Gender),"right"),X1$Item_Condition))
Y1$Vertical_Gender <- trimws(as.character(Y1$Vertical_Gender),"right")

X1 <- sqldf("SELECT bs.Vertical_Gender, bs.Item_Condition, X1.x
            FROM bs
            LEFT JOIN X1
            ON bs.New = X1.New")
X1$x[is.na(X1$x)] <- 0

Y1 <- sqldf("SELECT cs.Vertical_Gender, Y1.x
            FROM cs
            LEFT JOIN Y1
            ON cs.Vertical_Gender = Y1.Vertical_Gender")
Y1$x[is.na(Y1$x)] <- 0

# Replacing missing infor for Accessories - Handbags - Male for PA2
X2$New <- as.factor(paste(trimws(as.character(X2$Vertical_Gender),"right"),X2$Item_Condition))
Y2$Vertical_Gender <- trimws(as.character(Y2$Vertical_Gender),"right")

X2 <- sqldf("SELECT bs.Vertical_Gender, bs.Item_Condition, X2.x
            FROM bs
            LEFT JOIN X2
            ON bs.New = X2.New")
X2$x[is.na(X2$x)] <- 0

Y2 <- sqldf("SELECT cs.Vertical_Gender, Y2.x
            FROM cs
            LEFT JOIN Y2
            ON cs.Vertical_Gender = Y2.Vertical_Gender")
Y2$x[is.na(Y2$x)] <- 0

# Replacing missing infor for Accessories - Handbags - Male for VA2
B2$New <- as.factor(paste(trimws(as.character(B2$Vertical_Gender),"right"),B2$Item_Condition))
C2$Vertical_Gender <- trimws(as.character(C2$Vertical_Gender),"right")

B2 <- sqldf("SELECT bs.Vertical_Gender, bs.Item_Condition, B2.x
            FROM bs
            LEFT JOIN B2
            ON bs.New = B2.New")
B2$x[is.na(B2$x)] <- 0

C2 <- sqldf("SELECT cs.Vertical_Gender, C2.x
            FROM cs
            LEFT JOIN C2
            ON cs.Vertical_Gender = C2.Vertical_Gender")
C2$x[is.na(C2$x)] <- 0

setwd("C:\\Users\\vthangamuthu\\Documents\\Vijay\\Sam\\Authentication Survey\\Authentication Survey - V2\\Version 4\\UK")


write.csv(c(A1,D1,C1,B1),"UK_VA_1.csv")
write.csv(c(W1,Z1,Y1,X1),"UK_PA_1.csv")

write.csv(c(A2,D2,C2,B2),"UK_VA_2.csv")
write.csv(c(W2,Z2,Y2,X2),"UK_PA_2.csv")