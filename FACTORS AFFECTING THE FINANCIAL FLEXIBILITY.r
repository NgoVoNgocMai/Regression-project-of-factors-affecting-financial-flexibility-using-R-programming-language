# Cài đặt và tải các gói cần thiết
install.packages(c("readxl", "tidyr", "corrr", "fixest", "lmtest", "cowplot", "plm", "ggcorrplot", 
                   "imputeTS", "gridExtra", "corrplot", "GGally", "xtable", "knitr", "stargazer", 
                   "shiny", "ggplot2", "plotly", "dplyr","DT","psych","shiny","shinythemes"))
library(readxl)
library(tidyr)
library(corrr)
library(fixest)
library(lmtest)
library(cowplot)
library(plm)
library(ggcorrplot)
library(imputeTS)
library(gridExtra)
library(corrplot)
library(GGally)
library(xtable)
library(knitr)
library(stargazer)
library(shiny)
library(ggplot2)
library(plotly) 
library(dplyr)
library(DT)
library(psych)
library(shiny)
library(shinythemes)


#LÀM SẠCH VÀ CHUẨN BỊ DỮ LIỆU 
# Đọc dữ liệu từ các tệp Excel
df <- read_excel("K214142072.xlsx")

df2 <- df

# Đổi tên các cột để loại bỏ khoảng trắng hoặc thay bằng dấu gạch dưới
names(df2) <- gsub(" ", "_", names(df2))

# Đổi tên cột cho dễ sử dụng
df2 <- df2 %>%
  rename(Shareholders_Equity = `Shareholders'_Equity_-_Attributable_to_Parent_ShHold_-_Total`)
names(df2)
# Tính các chỉ số tài chính
df2 <- df2 %>%
  mutate(
    ROA = Net_Income_after_Tax / Total_Assets,
    CR = Total_Current_Assets / Total_Current_Liabilities,
    CH = `Cash_&_Cash_Equivalents` / Total_Assets,
    D_E = Total_Liabilities / Shareholders_Equity,
    CPI = CPI,
    SIZE = log(Total_Assets)
  )

# Tính toán Z-Score
df2 <- df2 %>%
  mutate(
    z_score = 1.2 * `Cash_&_Cash_Equivalents` / Total_Assets 
    +1.4 * `Retained_Earnings_-_Total` / Total_Assets 
    + 3.3 * `Earnings_before_Interest_&_Taxes_(EBIT)` / Total_Assets 
    +  0.6 * (Total_Assets - Total_Liabilities) / Total_Liabilities 
    +  0.99 * `Revenue_from_Business_Activities_-_Total` / Total_Assets,
    FF = `Cash_&_Cash_Equivalents` / Total_Assets + (1 - Total_Liabilities / Total_Assets)
  )

# Chọn các cột cần thiết
df_new <- df2[, c('code', 'date', 'ROA', 'CR', 'CH', 'D_E', 'CPI', 'SIZE', 'z_score', 'FF')]
df_new <- df_new[!duplicated(df_new[c("code", "date")], fromLast = TRUE), ]




# Hàm xử lý giá trị ngoại lai bằng IQR
remove_outliers <- function(x) {
  Q1 <- quantile(x, 0.25, na.rm = TRUE)
  Q3 <- quantile(x, 0.75, na.rm = TRUE)
  IQR <- Q3 - Q1
  x[x < (Q1 - 1.5 * IQR) | x > (Q3 + 1.5 * IQR)] <- NA
  return(x)
}

# Xử lý giá trị ngoại lai cho df1
df_new <- df_new %>%
  mutate(
    CR = remove_outliers(CR),
    D_E = remove_outliers(D_E),
    z_score = remove_outliers(z_score) )


#THỔNG KẾ MÔ TẢ DỮ LIỆU
summary(df_new)
describe(df_new)

df_new_desc <- describe(df_new)
kable(df_new_desc)
kable(df_new_desc, format = "latex")


cor_matrix <- cor(df_new[, c('ROA', 'CR', 'CH','D_E', 'CPI','SIZE','z_score','FF')], use = "complete.obs")
corrplot(cor_matrix, method = "circle", tl.cex = 0.8, number.cex = 0.8, addCoef.col = "black")
# Tạo bảng từ ma trận tương quan
cor_table <- xtable(cor_matrix)
# In bảng LaTeX
print(cor_table, include.rownames = TRUE)

# TRỰC  QUAN DƯ LIỆU 
# Biểu đồ histogram cho các biến
hist_ROA <- ggplot(df_new, aes(x = ROA)) +
  geom_histogram(binwidth = 0.01, fill = "skyblue", color = "black") +
  labs(title = "Distribution of ROA", x = "ROA", y = "Frequency")

hist_CR <- ggplot(df_new, aes(x = CR)) +
  geom_histogram(binwidth = 0.1, fill = "skyblue", color = "black") +
  labs(title = "Distribution of CR", x = "CR", y = "Frequency")

hist_CH <- ggplot(df_new, aes(x = CH)) +
  geom_histogram(binwidth = 0.02, fill = "skyblue", color = "black") +
  labs(title = "Distribution of CH", x = "CH", y = "Frequency")

hist_D_E <- ggplot(df_new, aes(x = D_E)) +
  geom_histogram(binwidth = 0.1, fill = "skyblue", color = "black") +
  labs(title = "Distribution of D_E", x = "D_E", y = "Frequency")

hist_cpi <- ggplot(df_new, aes(x = CPI)) +
  geom_histogram(binwidth = 0.05, fill = "skyblue", color = "black") +
  labs(title = "Distribution of CPI", x = "CPI", y = "Frequency")

hist_SIZE <- ggplot(df_new, aes(x = SIZE)) +
  geom_histogram(binwidth = 0.5, fill = "skyblue", color = "black") +
  labs(title = "Distribution of SIZE", x = "SIZE", y = "Frequency")

# Hiển thị tất cả các biểu đồ histogram trên cùng một mặt phẳng
grid.arrange(hist_ROA, hist_CR, hist_CH, hist_D_E, hist_cpi, hist_SIZE, ncol = 2)


# Biểu đồ boxplot cho các biến
boxplot_ROA <- ggplot(df_new, aes(y = ROA)) +
  geom_boxplot(fill = "skyblue", color = "black") +
  labs(title = "Boxplot of ROA", y = "ROA")

boxplot_CR <- ggplot(df_new, aes(y = CR)) +
  geom_boxplot(fill = "skyblue", color = "black") +
  labs(title = "Boxplot of CR", y = "CR")

boxplot_CH <- ggplot(df_new, aes(y = CH)) +
  geom_boxplot(fill = "skyblue", color = "black") +
  labs(title = "Boxplot of CH", y = "CH")

boxplot_D_E <- ggplot(df_new, aes(y = D_E)) +
  geom_boxplot(fill = "skyblue", color = "black") +
  labs(title = "Boxplot of D_E", y = "D_E")

boxplot_cpi <- ggplot(df_new, aes(y = CPI)) +
  geom_boxplot(fill = "skyblue", color = "black") +
  labs(title = "Boxplot of CPI", y = "CPI")

boxplot_SIZE <- ggplot(df_new, aes(y = SIZE)) +
  geom_boxplot(fill = "skyblue", color = "black") +
  labs(title = "Boxplot of SIZE", y = "SIZE")

# Hiển thị tất cả các biểu đồ boxplot trên cùng một mặt phẳng
grid.arrange(boxplot_ROA, boxplot_CR, boxplot_CH, boxplot_D_E, boxplot_cpi, boxplot_SIZE, ncol = 2)


#MÔ HÌNH HỒI QUY

#TRƯỚC COVID ( 2014 - 2019) - THANG ĐO BIẾN PHỤ THUỘC (FF)
df_2014_2019 <- df_new %>% filter(date >= 2014 & date <= 2019)
pdata <- plm::pdata.frame(df_2014_2019, index = c("code", "date"))
model1 <- plm(FF ~ ROA + CR + CH + D_E + CPI + SIZE, data = pdata, model = "within")
summary(model1)

#TRONG COVID ( 2020 - 2021)  - THANG ĐO BIẾN PHỤ THUỘC (FF)
df_2020_2021 <- df_new %>% filter(date >= 2020 & date <= 2021)
pdata <- plm::pdata.frame(df_2020_2021, index = c("code", "date"))
model2 <- plm(FF ~ ROA + CR + CH + D_E + CPI + SIZE, data = pdata, model = "within")
summary(model2)

#SAU COVID ( 2022 - 2023 )  - THANG ĐO BIẾN PHỤ THUỘC (FF)
df_2022_2023 <- df_new %>% filter(date >= 2022 & date <= 2023)
pdata <- plm::pdata.frame(df_2022_2023, index = c("code", "date"))
model3 <- plm(FF ~ ROA + CR + CH + D_E + CPI + SIZE, data = pdata, model = "within")
summary(model3)


#TRƯỚC COVID ( 2014 - 2019) - THANG ĐO BIẾN PHỤ THUỘC (z_score)
df_2014_2019 <- df_new %>% filter(date >= 2014 & date <= 2019)
pdata <- plm::pdata.frame(df_2014_2019, index = c("code", "date"))
model4 <- plm(z_score ~ ROA + CR + CH + D_E + CPI + SIZE, data = pdata, model = "within")
summary(model4)

#TRONG COVID ( 2020 - 2021)  - THANG ĐO BIẾN PHỤ THUỘC (z_score)
df_2020_2021 <- df_new %>% filter(date >= 2020 & date <= 2021)
pdata <- plm::pdata.frame(df_2020_2021, index = c("code", "date"))
model5 <- plm(z_score ~ ROA + CR + CH + D_E + CPI + SIZE, data = pdata, model = "within")
summary(model5)

#SAU COVID ( 2022 - 2023 )  - THANG ĐO BIẾN PHỤ THUỘC (z_score)
df_2022_2023 <- df_new %>% filter(date >= 2022 & date <= 2023)
pdata <- plm::pdata.frame(df_2022_2023, index = c("code", "date"))
model6 <- plm(z_score ~ ROA + CR + CH + D_E + CPI + SIZE, data = pdata, model = "within")
summary(model6)


# Combine all models in one LaTeX table
stargazer(model1, model2, model3, model4, model5, model6,
          type = "latex",
          title = "Regression Results",
          dep.var.labels = c("FF", "FF", "FF", "z_score", "z_score", "z_score"),
          covariate.labels = c("ROA", "CR", "CH", "DE", "CPI", "SIZE"),
          column.labels = c("Pre-COVID (2014-2019)", "During COVID (2020-2021)", "Post-COVID (2022-2023)",
                            "Pre-COVID (2014-2019)", "During COVID (2020-2021)", "Post-COVID (2022-2023)"),
          model.names = FALSE,
          out = "combined_regression_results.tex")
library(DT)
ui <- tagList(
  shinythemes::themeSelector(),
  navbarPage(
    "TÀI CHÍNH LINH HOẠT",
    
    tabPanel("TRỰC QUAN DỮ LIỆU",
             fluidRow(
               column(12, plotOutput(outputId = "histogram_plot"))
             ),
             fluidRow(
               column(12, plotOutput(outputId = "boxplot_plot"))
             )
    ),
    tabPanel("KẾT QUẢ MÔ HÌNH",
             sidebarPanel(
               selectInput("model_selection", "Chọn mô hình:",
                           choices = c("Mô hình trước Covid", "Mô hình trong Covid", "Mô hình sau Covid")),
               
             ), # Thêm dấu phẩy ở đây
             mainPanel(
               tableOutput("regression_table"),
               tableOutput("summary_table"),
               plotOutput(outputId = "correlation_plot")
             )
    ),
    tabPanel("DỮ LIỆU TÀI CHÍNH",
             sidebarPanel(
               # Thêm các điều khiển để chọn dữ liệu
               selectInput("financial_data", "Chọn dữ liệu:",
                           choices = c("Dữ liệu hồi quy trước Covid", "Dữ liệu hồi quy trong Covid", "Dữ liệu hồi quy sau Covid"),
                           
               ),
               mainPanel(
                 # Thêm đối tượng để hiển thị dữ liệu
                 DTOutput("selected_data_table")
               )
             )
    )
  )
)


# Define server function  
server <- function(input, output) {
  # Render continuous_plot
  output$histogram_plot <- renderPlot({
    # Biểu đồ histogram cho các biến
    hist_ROA <- ggplot(df_new, aes(x = ROA)) +
      geom_histogram(binwidth = 0.01, fill = "orange", color = "black") +
      labs(title = "Distribution of ROA", x = "ROA", y = "Frequency")
    
    hist_CR <- ggplot(df_new, aes(x = CR)) +
      geom_histogram(binwidth = 0.1, fill = "orange", color = "black") +
      labs(title = "Distribution of CR", x = "CR", y = "Frequency")
    
    hist_CH <- ggplot(df_new, aes(x = CH)) +
      geom_histogram(binwidth = 0.02, fill = "orange", color = "black") +
      labs(title = "Distribution of CH", x = "CH", y = "Frequency")
    
    hist_D_E <- ggplot(df_new, aes(x = D_E)) +
      geom_histogram(binwidth = 0.1, fill = "orange", color = "black") +
      labs(title = "Distribution of D_E", x = "D_E", y = "Frequency")
    
    hist_cpi <- ggplot(df_new, aes(x = CPI)) +
      geom_histogram(binwidth = 0.05, fill = "orange", color = "black") +
      labs(title = "Distribution of CPI", x = "CPI", y = "Frequency")
    
    hist_SIZE <- ggplot(df_new, aes(x = SIZE)) +
      geom_histogram(binwidth = 0.5, fill = "orange", color = "black") +
      labs(title = "Distribution of SIZE", x = "SIZE", y = "Frequency")
    
    # Hiển thị tất cả các biểu đồ histogram trên cùng một mặt phẳng
    grid.arrange(hist_ROA, hist_CR, hist_CH, hist_D_E, hist_cpi, hist_SIZE, ncol = 2)
  })
  
  # Render boxplot_plot
  output$boxplot_plot <- renderPlot({
    boxplot_ROA <- ggplot(df_new, aes(y = ROA)) +
      geom_boxplot(fill = "brown", color = "black") +
      labs(title = "Boxplot of ROA", y = "ROA")
    
    boxplot_CR <- ggplot(df_new, aes(y = CR)) +
      geom_boxplot(fill = "brown", color = "black") +
      labs(title = "Boxplot of CR", y = "CR")
    
    boxplot_CH <- ggplot(df_new, aes(y = CH)) +
      geom_boxplot(fill = "brown", color = "black") +
      labs(title = "Boxplot of CH", y = "CH")
    
    boxplot_D_E <- ggplot(df_new, aes(y = D_E)) +
      geom_boxplot(fill = "brown", color = "black") +
      labs(title = "Boxplot of D_E", y = "D_E")
    
    boxplot_cpi <- ggplot(df_new, aes(y = CPI)) +
      geom_boxplot(fill = "brown", color = "black") +
      labs(title = "Boxplot of CPI", y = "CPI")
    
    boxplot_SIZE <- ggplot(df_new, aes(y = SIZE)) +
      geom_boxplot(fill = "brown", color = "black") +
      labs(title = "Boxplot of SIZE", y = "SIZE")
    
    # Hiển thị tất cả các biểu đồ boxplot trên cùng một mặt phẳng
    grid.arrange(boxplot_ROA, boxplot_CR, boxplot_CH, boxplot_D_E, boxplot_cpi, boxplot_SIZE, ncol = 2)
  })
  
  
  # Render correlation_plot
  output$correlation_plot <- renderPlot({
    cor_matrix <- cor(df_new[, c('ROA', 'CR', 'CH','D_E', 'CPI','SIZE','z_score','FF')], use = "complete.obs")
    corrplot(cor_matrix, method = "circle", tl.cex = 0.8, number.cex = 0.8, addCoef.col = "black")
  })
  
  # Render summary_table
  output$summary_table <- renderTable({
    describe(df_new[, c('ROA', 'CR', 'CH','D_E', 'CPI','SIZE','z_score','FF')])
  })
  
  
  # Render regression_table
  output$regression_table <- renderTable({
    model_selection <- switch(input$model_selection,
                              "Mô hình trước Covid" = plm(FF ~ ROA + CR + CH  + D_E + CPI + SIZE, data = df_new[df_new$date < 2020, ], model = "within"),
                              "Mô hình trong Covid" = plm(FF ~ ROA + CR + CH  + D_E + CPI + SIZE, data = df_new[df_new$date >= 2020 & df_new$date <= 2021, ], model = "within"),
                              "Mô hình sau Covid" = plm(FF ~ ROA + CR + CH  + D_E + CPI + SIZE, data = df_new[df_new$date > 2021, ], df_new = "within")
    )
    
    model_summary <- summary(model_selection)
    model_table <- as.data.frame(model_summary$coefficients)
    model_table$`Pr(>|t|)` <- model_summary$coefficients[, "Pr(>|t|)"]
    
    # Add a column for independent variables
    model_table <- tibble::rownames_to_column(model_table, "Biến độc lập")
    
    # Reorder columns
    model_table <- model_table[, c("Biến độc lập", "Estimate", "Std. Error", "t-value", "Pr(>|t|)")]
  })
  
  
  
  
  # Khai báo một reactive chứa dữ liệu đã chọn
  selected_data <- reactive({
    switch(input$financial_data,
           "Dữ liệu hồi quy trước Covid" = df_new[df_new$date < 2020, ],
           "Dữ liệu hồi quy trong Covid" = df_new[df_new$date >= 2020 & df_new$date <= 2021, ],
           "Dữ liệu hồi quy sau Covid" = df_new[df_new$date > 2021, ]
    )
  })
  
  # Hiển thị dữ liệu đã chọn
  output$selected_data_table <- renderDT({
    datatable(selected_data())
  })
  
}

# Create Shiny app 
shinyApp(ui = ui, server = server)


