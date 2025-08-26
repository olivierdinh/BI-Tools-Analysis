# Import Excel sheet
library(readxl)
df <- read_excel("C:/Users/dinho2/Medtronic PLC/Team - Thesis - Thesis/Thesis/4. References/BI Tools.xlsx", sheet = "4. Top 30 - R")

# Core scoring formulas
df$"Review × Ratings" <- round(df$Review * df[["Ratings Number"]], 0)                         # Review × Ratings
df$"Review × √Ratings" <- round(df$Review * sqrt(df[["Ratings Number"]]), 0)                  # Review × √Ratings
df$"Review × log(Ratings)" <- round(df$Review * log(df[["Ratings Number"]]), 0)               # Review × log(Ratings)
df$"Review × Ratings^0.5" <- round(df$Review * df[["Ratings Number"]]^0.5, 0)                 # Review × Ratings²
df$"Review² × Ratings" <- round(df$Review^2 * df[["Ratings Number"]], 0)                      # Review² × Ratings
df$"Review² × √Ratings" <- round(df$Review^2 * sqrt(df[["Ratings Number"]]), 0)               # Review² × √Ratings
df$"Review² × log(Ratings)" <- round(df$Review^2 * log(df[["Ratings Number"]]), 0)            # Review² × log(Ratings)
df$"Review² × Ratings^0.5" <- round(df$Review^2 * (df[["Ratings Number"]]^0.5), 0)            # Review² × Ratings^0.5

# Rankings
df$"Rank: Review × Ratings" <- rank(-df$"Review × Ratings")
df$"Rank: Review × √Ratings" <- rank(-df$"Review × √Ratings")
df$"Rank: Review × log(Ratings)" <- rank(-df$"Review × log(Ratings)")
df$"Rank: Review × Ratings^0.5" <- rank(-df$"Review × Ratings^0.5")
df$"Rank: Review² × Ratings" <- rank(-df$"Review² × Ratings")
df$"Rank: Review² × √Ratings" <- rank(-df$"Review² × √Ratings")
df$"Rank: Review² × log(Ratings)" <- rank(-df$"Review² × log(Ratings)")
df$"Rank: Review² × Ratings^0.5" <- rank(-df$"Review² × Ratings^0.5")


library(dplyr)
# Rank: Review × Ratings (with tie-breaker)
df <- df %>%
  arrange(desc(`Review × Ratings`), desc(`Ratings Number`)) %>%
  mutate(`Rank: Review × Ratings` = row_number())

# Rank: Review × √Ratings
df <- df %>%
  arrange(desc(`Review × √Ratings`), desc(`Ratings Number`)) %>%
  mutate(`Rank: Review × √Ratings` = row_number())

# Rank: Review × log(Ratings)
df <- df %>%
  arrange(desc(`Review × log(Ratings)`), desc(`Ratings Number`)) %>%
  mutate(`Rank: Review × log(Ratings)` = row_number())

# Rank: Review × Ratings^0.5
df <- df %>%
  arrange(desc(`Review × Ratings^0.5`), desc(`Ratings Number`)) %>%
  mutate(`Rank: Review × Ratings^0.5` = row_number())

# Rank: Review² × Ratings
df <- df %>%
  arrange(desc(`Review² × Ratings`), desc(`Ratings Number`)) %>%
  mutate(`Rank: Review² × Ratings` = row_number())

# Rank: Review² × √Ratings
df <- df %>%
  arrange(desc(`Review² × √Ratings`), desc(`Ratings Number`)) %>%
  mutate(`Rank: Review² × √Ratings` = row_number())

# Rank: Review² × log(Ratings)
df <- df %>%
  arrange(desc(`Review² × log(Ratings)`), desc(`Ratings Number`)) %>%
  mutate(`Rank: Review² × log(Ratings)` = row_number())

# Rank: Review² × Ratings^0.5
df <- df %>%
  arrange(desc(`Review² × Ratings^0.5`), desc(`Ratings Number`)) %>%
  mutate(`Rank: Review² × Ratings^0.5` = row_number())


library(ggplot2)
# Compare ranks across scoring systems for a few top tools
top_tools <- df[df$Tool %in% c("Tableau (+ CRM Analytics)", "Microsoft Power BI", "Qlik (View + Sense + Cloud Analytics)", "Looker", "Databricks Data Intelligence Platform"), ]

# Melt data to long format if you want line plots per tool
library(tidyr)
ranks_long <- pivot_longer(top_tools, starts_with("rank"), names_to = "Method", values_to = "Rank")

ggplot(ranks_long, aes(x = Method, y = Rank, group = Tool, color = Tool)) +
  geom_line() + geom_point() +
  scale_y_reverse() +
  labs(title = "Ranking Across Scoring Systems", y = "Rank (Lower is Better)", x = "Scoring Method") +
  theme_minimal()





library(writexl)
write_xlsx(df, "BI_Tools_Scoring_Comparison.xlsx")
