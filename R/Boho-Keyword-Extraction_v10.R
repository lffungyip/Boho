# ####Understanding and Testing #### 
# library(dplyr)
# 
# My_Items <- readRDS("TA_Items-JimPAmazonBoho.Rds")  %>%
#   filter(Sheet_ID == 4) 
# 
# My_Style <- readRDS("TA_Style_WordsJimPAmazonBoho.Rds")  %>%
#   filter(Sheet_ID == 4) 
# 
# My_Disc <- readRDS("TA_Disc_WordsJimPAmazonBoho.Rds")  %>%
#   filter(Sheet_ID == 4) 

#################################################################

# install.packages("readxl")
library(readxl)
# install.packages("readr")
library(readr)
library(dplyr)
# install.packages("tidyr")
library(tidyr)
library(tibble)
library(stringr)
library(tm)
# install.packages("uuid")
library(uuid)

Proj_Name <- "JimPAmazonBoho"
# Data_Path <- "../DataIn/Trendalytics"
Data_Path <- "C:/Users/FungYip/Documents/AmazonBoho-master/DataIn/Trendalytics"
getwd()

####Load Source Sheets to TA_Items####
sheets <- tibble(Sheet_Name = dir(Data_Path)) %>% 
  mutate(Modified_Date = file.mtime(paste0(Data_Path, "/", Sheet_Name)),
         Sheet_ID = row_number()) %>% 
  select(Sheet_ID, Sheet_Name, Modified_Date)
TA_Items <- NULL

x<-c(1,2,3,4,5)

for(i in seq_along(x)){
  items_temp <- read_excel(paste0(Data_Path, "/", sheets$Sheet_Name[i]), skip = 1)
  items_temp <- cbind(items_temp, sheets[i,]) %>% 
    mutate(Item_ID = row_number())
  TA_Items <- rbind(TA_Items, items_temp)
}

TA_Items <- as_tibble(TA_Items) %>% 
  mutate(extract_Date = as.Date(extractDate)) %>% 
  select(Sheet_ID, Sheet_Name, Modified_Date, Item_ID, Category, Brand, Retailer, 
         `Style Name`, Color, Currency, `Original Price`, `Markdown (%)`,
         extractDate, `Product Description`) %>% 
  arrange(Sheet_ID, Item_ID)

TA_Items <- TA_Items %>%    
  filter(!(is.na(TA_Items$Category) | TA_Items$Category==""))

#### Break out words in Style Name and Product Description to TA_Style_Words and TA_Desc_Words respectively####
max_num_words <- max(str_count(TA_Items$`Style Name`, " "), na.rm = TRUE) + 1
cat("max_num_words: ", max_num_words, "\n")

#< vectorized opps & tidy it
ss <- as_tibble(str_split_fixed(TA_Items$`Style Name`, " ", n = max_num_words))
ss$Item_ID <- TA_Items$Item_ID
ss$Sheet_ID <- TA_Items$Sheet_ID
TA_Style_Words <- ss %>%
  gather(Word_Position, Word, -c(Sheet_ID, Item_ID)) %>%   
  mutate(Word = tolower(str_replace_all(Word, "[^a-zA-Z0-9]", "")),
         Word_Position = as.integer(str_replace(Word_Position, "^V", ""))) %>%
  filter(Word != "") %>% 
  select(Sheet_ID, Item_ID, Word_Position, Word) %>% 
  arrange(Sheet_ID, Item_ID, Word_Position)
TA_Style_Words

## Description
max_num_words <- max(str_count(TA_Items$`Product Description`, " "), na.rm = TRUE) + 1
cat("max_num_words: ", max_num_words, "\n")


#< vectorized opps & tidy it
ss <- as_tibble(str_split_fixed(TA_Items$`Product Description`, " ", n = max_num_words))
ss$Item_ID <- TA_Items$Item_ID
ss$Sheet_ID <- TA_Items$Sheet_ID
TA_Disc_Words <- ss %>%
  gather(Word_Position, Word, -c(Sheet_ID, Item_ID)) %>%   
  mutate(Word = tolower(str_replace_all(Word, "[^a-zA-Z0-9]", "")),
         Word_Position = as.integer(str_replace(Word_Position, "^V", ""))) %>%
  filter(Word != "") %>% 
  select(Sheet_ID, Item_ID, Word_Position, Word) %>% 
  arrange(Sheet_ID, Item_ID, Word_Position)
TA_Disc_Words


#### Remove stopwords. Then some EDA on word frequency####
swds <- as_tibble(stopwords())
colnames(swds) <- "Word"
TA_Style_Words_Total_Count <- nrow(TA_Style_Words)
TA_Style_Words <- anti_join(TA_Style_Words, swds)

TA_Style_Words_NoStop_Count <- nrow(TA_Style_Words)
TA_Disc_Words_Total_Count <- nrow(TA_Disc_Words)
TA_Disc_Words <- anti_join(TA_Disc_Words, swds)

TA_Disc_Words_NoStop_Count <- nrow(TA_Disc_Words)

#Style
TA_Style_Frequency <- TA_Style_Words %>% 
  group_by(Word) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
TA_Style_Frequency

#Desc
TA_Desc_Frequency <- TA_Disc_Words %>% 
  group_by(Word) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
TA_Desc_Frequency

#Output
saveRDS(TA_Items, paste0("TA_Items-", Proj_Name, ".Rds"))
saveRDS(TA_Style_Words, paste0("TA_Style_Words", Proj_Name, ".Rds"))
saveRDS(TA_Disc_Words, paste0("TA_Disc_Words", Proj_Name, ".Rds"))



TA_Style_Words_v1 <- TA_Style_Words %>% mutate(RevWord = toupper(Word))
TA_Disc_Words_v1 <- TA_Disc_Words %>% mutate(RevWord = toupper(Word))

TA_Style_Disc_Words<-rbind(TA_Style_Words_v1,TA_Disc_Words_v1)
#Distinct
TA_Style_Disc_Words_d<-TA_Style_Disc_Words %>% distinct(Sheet_ID,Item_ID,RevWord)

####Matching####
getwd()
mylibrary <- read_excel("Product_Classification_Raw.xlsx") 

#Final Output
MyWords_mylibrary<-TA_Style_Disc_Words %>% inner_join(mylibrary,by="RevWord")

# MyWords_mylibrary_d<-TA_Style_Disc_Words_d %>% inner_join(mylibrary,by="RevWord")

saveRDS(MyWords_mylibrary, paste0("MyWords_mylibrary.Rds"))
write.csv(MyWords_mylibrary,"MyWords_mylibrary.csv")

####Run 3####
#Question 1 Tunic
MyWords_mylibrary_Part<- MyWords_mylibrary %>% 
  filter(RevWord=="TUNIC") %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Sheet_ID,Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Part


offshoulder <- c("OFF THE SHOULDER","OFFSHOULDER","OFFTHESHOULDER","OFF-THE-SHOULDER")
MyWords_mylibrary_Part<- MyWords_mylibrary %>% 
  filter(RevWord %in% offshoulder) %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Part

coldshoulder <- c("COLD SHOULDER","COLDSHOULDER")
MyWords_mylibrary_Part<- MyWords_mylibrary %>% 
  filter(RevWord %in% coldshoulder) %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Part

library(xlsx)
library(ggplot2)


####by Pattern Attribute####
MyWords_mylibrary_Pattern<- MyWords_mylibrary %>% 
  filter(SubTag=="Pattern") %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Pattern

write.xlsx(MyWords_mylibrary_Pattern,"MyWords_mylibrary_Pattern.xlsx")

MyWords_mylibrary_Pattern$label <- paste(MyWords_mylibrary_Pattern$RevWord,
                                        MyWords_mylibrary_Pattern$Number_Items_Containing, sep = ", ")
treemap(MyWords_mylibrary_Pattern,
        index=c("label"),
        vSize="Number_Items_Containing",
        vColor="Number_Items_Containing",
        type="value")

MyWords_mylibrary_Pattern_d<- MyWords_mylibrary %>% 
  filter(SubTag=="Pattern") %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Sheet_ID,Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Pattern_d

write.xlsx(MyWords_mylibrary_Pattern,"MyWords_mylibrary_Pattern.xlsx")

MyWords_mylibrary_Pattern$label <- paste(MyWords_mylibrary_Pattern$RevWord,
                                         MyWords_mylibrary_Pattern$Number_Items_Containing, sep = ", ")
treemap(MyWords_mylibrary_Pattern,
        index=c("label"),
        vSize="Number_Items_Containing",
        vColor="Number_Items_Containing",
        type="value")

####by Material Attribute####
MyWords_mylibrary_Material<- MyWords_mylibrary %>% 
  filter(SubTag=="Material") %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Material

write.xlsx(MyWords_mylibrary_Material,"MyWords_mylibrary_Material.xlsx")

MyWords_mylibrary_Material$label <- paste(MyWords_mylibrary_Material$RevWord,
                                         MyWords_mylibrary_Material$Number_Items_Containing, sep = ", ")
treemap(MyWords_mylibrary_Material,
        index=c("label"),
        vSize="Number_Items_Containing",
        vColor="Number_Items_Containing",
        type="value")

####by Structure Attribute####
MyWords_mylibrary_Structure<- MyWords_mylibrary %>% 
  filter(SubTag=="Structure") %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Structure

write.xlsx(MyWords_mylibrary_Structure,"MyWords_mylibrary_Structure.xlsx")

MyWords_mylibrary_Structure$label <- paste(MyWords_mylibrary_Structure$RevWord,
                                          MyWords_mylibrary_Structure$Number_Items_Containing, sep = ", ")
treemap(MyWords_mylibrary_Structure,
        index=c("label"),
        vSize="Number_Items_Containing",
        vColor="Number_Items_Containing",
        type="value")

# ####by Look Attribute####
# MyWords_mylibrary_Look<- MyWords_mylibrary %>% 
#   filter(SubTag=="Look") %>% 
#   group_by(RevWord) %>% 
#   summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
#   arrange(desc(Number_Items_Containing))
# MyWords_mylibrary_Look
# 
# MyWords_mylibrary_Look$label <- paste(MyWords_mylibrary_Look$RevWord,
#                                            MyWords_mylibrary_Look$Number_Items_Containing, sep = ", ")
# treemap(MyWords_mylibrary_Look,
#         index=c("label"),
#         vSize="Number_Items_Containing",
#         vColor="Number_Items_Containing",
#         type="value")

# ####by Person Attribute####
# MyWords_mylibrary_Person<- MyWords_mylibrary %>% 
#   filter(SubTag=="Person") %>% 
#   group_by(RevWord) %>% 
#   summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
#   arrange(desc(Number_Items_Containing))
# MyWords_mylibrary_Person

####by Style Attribute####
MyWords_mylibrary_Style<- MyWords_mylibrary %>% 
  filter(SubTag=="Style") %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Style

MyWords_mylibrary_Style$label <- paste(MyWords_mylibrary_Style$RevWord,
                                           MyWords_mylibrary_Style$Number_Items_Containing, sep = ", ")

write.xlsx(MyWords_mylibrary_Style,"MyWords_mylibrary_Style.xlsx")


treemap(MyWords_mylibrary_Style,
        index=c("label"),
        vSize="Number_Items_Containing",
        vColor="Number_Items_Containing",
        type="value")


####by Shape Attribute####
MyWords_mylibrary_Shape<- MyWords_mylibrary %>% 
  filter(SubTag=="Shape") %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Shape

MyWords_mylibrary_Shape$label <- paste(MyWords_mylibrary_Shape$RevWord,
                                       MyWords_mylibrary_Shape$Number_Items_Containing, sep = ", ")

write.xlsx(MyWords_mylibrary_Shape,"MyWords_mylibrary_Shape.xlsx")


treemap(MyWords_mylibrary_Shape,
        index=c("label"),
        vSize="Number_Items_Containing",
        vColor="Number_Items_Containing",
        type="value")

####by Part Attribute####
MyWords_mylibrary_Part<- MyWords_mylibrary %>% 
  filter(SubTag=="Part") %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(paste(Sheet_ID,Item_ID))) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Part

write.xlsx(MyWords_mylibrary_Part,"MyWords_mylibrary_Part_new.xlsx")


MyWords_mylibrary_Part$label <- paste(MyWords_mylibrary_Part$RevWord,
                                       MyWords_mylibrary_Part$Number_Items_Containing, sep = ", ")
treemap(MyWords_mylibrary_Part,
        index=c("label"),
        vSize="Number_Items_Containing",
        vColor="Number_Items_Containing",
        type="value")









####Run 2####
####by Attributes####
MyWords_mylibrary_Attributes <- MyWords_mylibrary %>%
  filter(MainTag=="Attributes") %>%
  filter(Sheet_ID=='5') %>% 
  group_by(SubTag) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Attributes

####by Colour Attribute####
MyWords_mylibrary_Colour<- MyWords_mylibrary %>% 
  filter(SubTag=="Colour") %>% 
  filter(Sheet_ID=='5') %>% 
    group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Colour

####by Pattern Attribute####
MyWords_mylibrary_Pattern<- MyWords_mylibrary %>% 
  filter(SubTag=="Pattern") %>%
  filter(Sheet_ID=='5') %>% 
    group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Pattern

####by Material Attribute####
MyWords_mylibrary_Material<- MyWords_mylibrary %>% 
  filter(SubTag=="Material") %>%
  filter(Sheet_ID=='5') %>% 
    group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Material

####by Structure Attribute####
MyWords_mylibrary_Structure<- MyWords_mylibrary %>% 
  filter(SubTag=="Structure") %>%
  filter(Sheet_ID=='5') %>% 
    group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Structure

####by Look Attribute####
MyWords_mylibrary_Look<- MyWords_mylibrary %>% 
  filter(SubTag=="Look") %>% 
  filter(Sheet_ID=='5') %>% 
    group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Look

####by Person Attribute####
MyWords_mylibrary_Person<- MyWords_mylibrary %>% 
  filter(SubTag=="Person") %>%
  filter(Sheet_ID=='5') %>% 
    group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Person

####by Style Attribute####
MyWords_mylibrary_Style<- MyWords_mylibrary %>% 
  filter(SubTag=="Style") %>%
  filter(Sheet_ID=='5') %>% 
    group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Style

####by Shape Attribute####
MyWords_mylibrary_Shape<- MyWords_mylibrary %>% 
  filter(SubTag=="Shape") %>%
  filter(Sheet_ID=='5') %>% 
    group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Shape

####by Part Attribute####
MyWords_mylibrary_Part<- MyWords_mylibrary %>% 
  filter(SubTag=="Part") %>%
  filter(Sheet_ID=='5') %>% 
    group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Part


MyWords_mylibrary_Attributes <- MyWords_mylibrary %>%
  filter(MainTag=="Attributes") %>%
  filter(Sheet_ID=='5') %>% 
  group_by(SubTag) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Attributes

####by Colour Attribute####
MyWords_mylibrary_Colour<- MyWords_mylibrary %>% 
  filter(SubTag=="Colour") %>% 
  filter(Sheet_ID=='5') %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Colour

####by Pattern Attribute####
MyWords_mylibrary_Pattern<- MyWords_mylibrary %>% 
  filter(SubTag=="Pattern") %>%
  filter(Sheet_ID=='5') %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Pattern

####by Material Attribute####
MyWords_mylibrary_Material<- MyWords_mylibrary %>% 
  filter(SubTag=="Material") %>%
  filter(Sheet_ID=='5') %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Material

####by Structure Attribute####
MyWords_mylibrary_Structure<- MyWords_mylibrary %>% 
  filter(SubTag=="Structure") %>%
  filter(Sheet_ID=='5') %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Structure

####by Look Attribute####
MyWords_mylibrary_Look<- MyWords_mylibrary %>% 
  filter(SubTag=="Look") %>% 
  filter(Sheet_ID=='5') %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Look

####by Person Attribute####
MyWords_mylibrary_Person<- MyWords_mylibrary %>% 
  filter(SubTag=="Person") %>%
  filter(Sheet_ID=='5') %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Person

####by Style Attribute####
MyWords_mylibrary_Style<- MyWords_mylibrary %>% 
  filter(SubTag=="Style") %>%
  filter(Sheet_ID=='5') %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Style

####by Shape Attribute####
MyWords_mylibrary_Shape<- MyWords_mylibrary %>% 
  filter(SubTag=="Shape") %>%
  filter(Sheet_ID=='5') %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Shape

####by Part Attribute####
MyWords_mylibrary_Part<- MyWords_mylibrary %>% 
  filter(SubTag=="Part") %>%
  filter(Sheet_ID=='5') %>% 
  group_by(RevWord) %>% 
  summarise(Number_Items_Containing = n_distinct(Item_ID)) %>% 
  arrange(desc(Number_Items_Containing))
MyWords_mylibrary_Part


# install.packages("treemap")
library(treemap)
MyWords_mylibrary_Part$label <- paste(MyWords_mylibrary_Part$RevWord,
                                      MyWords_mylibrary_Part$Number_Items_Containing, sep = ", ")
treemap(MyWords_mylibrary_Part,
        index=c("label"),
        vSize="Number_Items_Containing",
        vColor="Number_Items_Containing",
        type="value")

