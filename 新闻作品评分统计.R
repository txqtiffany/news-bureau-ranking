# Install required libraries if needed
install.packages("readxl")
install.packages("openxlsx")
install.packages("dplyr")
install.packages("tidyr")
install.packages("magrittr")
install.packages("purrr")

# Load libraries
library(readxl)
library(openxlsx)
library(dplyr)
library(tidyr)
library(magrittr)
library(purrr)

# Get a directory list of all files in the folder
# 读取全部文件路径
社级文件路径 <- list.files(path = "./2023年好稿/社级好稿/",
                     pattern = ".xls$", full.names = TRUE)
总编室文件路径 <- list.files(path = "./2023年好稿/总编室好稿/",
                      pattern = ".xls$", full.names = TRUE)
部级文件路径 <- list.files(path = "./2023年好稿/部级好稿/",
                     pattern = ".xls$", full.names = TRUE)

# Organize all articles within the same level into one sheet and mark its origin
# 以好稿级别汇总稿件并标注来源
社级文件汇总 <- 社级文件路径 %>%
  map_dfr(~{
    文件路径 <- .x
    read_xls(文件路径) %>%
      select("稿件标题", "作者", "编辑") %>%
      mutate(稿件级别 = "社级",
             来源 = basename(文件路径))
})

总编室文件汇总 <- 总编室文件路径 %>%
  map_dfr(~{
    文件路径 <- .x
    read_xls(文件路径) %>%
      select("稿件标题", "作者", "编辑") %>%
      mutate(稿件级别 = "总编室",
             来源 = basename(文件路径))
  })

部级文件汇总 <- 部级文件路径 %>%
  map_dfr(~{
    文件路径 <- .x
    read_xls(文件路径) %>%
      select("稿件标题", "作者", "编辑") %>%
      mutate(稿件级别 = "部级",
             来源 = basename(文件路径))
  })


好稿汇总 <- 社级文件汇总 %>%
  bind_rows(总编室文件汇总) %>%
  bind_rows(部级文件汇总)


# Get the list of EU bureaus 
# 读取欧洲分社名单
欧洲分社list <- read.xlsx("./2023年欧洲地区好稿统计（上半年）.xlsx",
                      sheet = 1, startRow = 2, ) %>%
  select("分社名称") %>%
  slice(-n())  # dropping the last row of "总计"
  
欧洲分社list <- 欧洲分社list[[1]]

# Add a new column to indicate bureau presence
# 添加新的列代表稿件作者所属社
好稿汇总$作者所属分社 <- NA
好稿汇总$编辑所属分社 <- NA

# Iterate through each bureau
# 迭代欧洲地区全部分社
for (分社 in 欧洲分社list) {
  # Check if the bureau appears in the author or editor column
  # 为作者或编辑栏出现欧洲地区分社的稿件添加标签
  好稿汇总 %<>%
    mutate(作者所属分社 = ifelse(grepl(分社, 作者),
                           ifelse(is.na(作者所属分社), 分社,
                                  paste(作者所属分社, 分社, sep = ", ")),
                           作者所属分社)) %>%
    mutate(编辑所属分社 = ifelse(grepl(分社, 编辑), 
                           ifelse(is.na(编辑所属分社), 分社,
                                  paste(编辑所属分社, 分社, sep = ", ")),
                           编辑所属分社))
}

# Remove articles that are not from EU bureaus
# 筛选出欧洲地区的稿件
欧洲地区好稿汇总 <- 好稿汇总 %>%
  filter(!is.na(作者所属分社) | !is.na(编辑所属分社)) %>%
  select("稿件级别", "稿件标题", "作者所属分社", "编辑所属分社", "来源")

# Rearrange into bureaus and their respective articles
# 以分社为单位整理其稿件汇总
作者类好稿汇总 <- 欧洲地区好稿汇总 %>%
  select("稿件级别", "作者所属分社", "稿件标题", "来源") %>%
  rename(分社 = 作者所属分社) %>%
  filter(!is.na(分社)) %>%
  separate_rows(分社, sep = ", ") %>%
  arrange(分社)

编辑类好稿汇总 <- 欧洲地区好稿汇总 %>%
  select("稿件级别", "编辑所属分社", "稿件标题", "来源") %>%
  rename(分社 = 编辑所属分社) %>%
  filter(!is.na(分社)) %>%
  separate_rows(分社, sep = ", ") %>%
  arrange(分社)

# Rank bureaus base on rating system
# 以分社为单位进行评分
分社作者类好稿评分 <- 作者类好稿汇总 %>%
  group_by(分社, 稿件级别) %>%
  summarise(稿件数量 = n()) %>%
  pivot_wider(names_from = 稿件级别, values_from = 稿件数量, values_fill = 0) %>%
  mutate(评分 = (社级 * 4 + 总编室 * 2 + 部级 * 1)) %>%
  select("分社", "社级", "总编室", "部级", "评分") %>%
  arrange(-评分)

分社编辑类好稿评分 <- 编辑类好稿汇总 %>%
  group_by(分社, 稿件级别) %>%
  summarise(稿件数量 = n()) %>%
  pivot_wider(names_from = 稿件级别, values_from = 稿件数量, values_fill = 0) %>%
  mutate(评分 = (社级 * 4 + 总编室 * 2 + 部级 * 1)) %>%
  select("分社", "社级", "总编室", "部级", "评分") %>%
  arrange(-评分)


# Generate output into one excel file 
# 在同一路径下导出表格
write.xlsx(list(作者类好稿数量统计 = 分社作者类好稿评分,
                编辑类好稿数量统计 = 分社编辑类好稿评分,
                作者类好稿汇总 = 作者类好稿汇总,
                编辑类好稿汇总 = 编辑类好稿汇总,
                欧洲总分社好稿汇总 = 欧洲地区好稿汇总), 
           "./2023年欧洲地区好稿统计.xlsx")
