dt <- mtcars


?barplot
barplot(height = table(dt$am), col = "red")

class(dt$am)
dt$am <- as.factor(dt$am)
class(dt$am)

levels(dt$am)
levels(dt$am) <- 
  c(
    "0" = "low", "1" = "high"
  )

barplot(height = table(dt$am), col = "red")


barplot(height = table(dt$am), 
        col = "#43ADE3",
        main = "main",
        xlab = "xlab",
        ylab = "ylab",
        ylim = c(0, 22)
)

hist(dt$wt, 
     col = "#43ADE3",
     main = "main",
     xlab = "xlab",
     ylab = "ylab",
     xlim = c(1, 6)
)

hist(dt$wt, 
     col = "#43ADE3",
     main = "main",
     xlab = "xlab",
     ylab = "ylab",
     xlim = c(1, 6),
     ylim = c(0, 10),
     labels = TRUE
)


gss2014 <- 
  read.csv("/Users/ophirbetser/Ophir/Teaching/katya/GSS2014.csv")


mean(gss2014$HRS1, na.rm = TRUE)
?mean


quantile(gss2014$REALINC, 
         probs = c(0.1, 0.2, 0.3, 0.4),
         na.rm = TRUE)




boxplot(gss2014$REALINC,
         xlab = "Incomes",
         border = "red",
         frame.plot = TRUE
)

boxplot(gss2014$REALINC,
        xlab = "Incomes",
        border = "red",
        frame.plot = FALSE
)



boxplot(gss2014$REALINC ~ gss2014$DIVORCE)








