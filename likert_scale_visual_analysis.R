library(readxl)
library(likert)
library(magrittr)


# Configs
orig.data.path <- "./tables"
file.names <- list.files(orig.data.path)
file.paths <- c()
for (f in file.names) {
  file.paths <- c(file.paths, paste0(orig.data.path, "/", f))
}

table.nums <- stringr::str_extract_all(file.paths, "\\d+", simplify = TRUE) %>%
  as.integer() %>%
  sort()

path.to.save <- paste0("./", strsplit(orig.data.path, "[.]")[[1]][1])
smpl.size.ls <- list(
   "1" = 101,
   "2" = 101,
   "6" = 101,
  "20" = 101,
  "22" =  43,
  "23" =  26,
  "24" =  32,
  "25" = 101,
  "28" = 101,
  "39" = 101,
  "40" = 101
)
# N <- 101 # Sample Size of original survey
quest.col.name <- "Հարց"

lang <- "pol"
sheet <- if (lang == "pol") 1 else 2

# for (i in 1:length(file.paths)) {
# for (i in table.nums)
for (i in as.numeric(names(smpl.size.ls)))
{ #i = 2
  N <- smpl.size.ls[[as.character(i)]]
  # Reading and pre-processing data
  # data. <- read_excel(file.paths[i], sheet = sheet)
  filepath <- paste0(
    orig.data.path, "/",
    "Table ", i, ".xlsx"
  )
  
  data. <- read_excel(filepath, sheet = if (i == 2) 2 else sheet)
  colnames(data.)[1] <- quest.col.name
  
  data.num <- data.[2:dim(data.)[1], c(1, seq(2,dim(data.)[2], 2))]
  colnames(data.num)[1] <- quest.col.name
  
  data.prct <- data.[2:dim(data.)[1], seq(1,dim(data.)[2], 2)]
  colnames(data.prct)[1] <- quest.col.name
  
  # Use in case, there is a need for custom colnames
  # cols <- c(
  #   "Հարց", "Մեծ", "Փոքր",
  #   "Գրեթե ոչ", "Ոչ մի",
  #   "Դժվարանում եմ"
  # )
  
  cols <- colnames(data.num)
  
  colnames(data.prct) <- cols
  
  # Generating synthetic survey data
  data.survey <- data.frame(ID = 1:N)
  
  for (quest in data.num[[quest.col.name]]) {
    quest.row = data.num[data.num[quest.col.name] == quest, ]
    # quest.row = data.[data.[quest.col.name] == quest, ]
    surv.vect = c()
    for (j in 2:length(cols)) {
      num = quest.row[ , j][[1]]
      surv.vect = c(
        surv.vect,
        rep(names(quest.row[ , j]), as.integer(num))
      )
    }
    data.survey[quest] <- factor(surv.vect, levels=cols[-1])
  }
  
  # Getting Likert matrix
  prct_matrix <- likert(
    items = data.survey[,2:dim(data.survey)[2]], 
    nlevels = 5
  )
  
  # Generating Plots
  
  path.to.save <- paste0(
    path.to.save, 
    "graphs/", 
    strsplit(
      strsplit(filepath, "/")[[1]][3], # file.paths[i]
      "\\."
    )[[1]][1]
  )
  
  if (!dir.exists(path.to.save))
    dir.create(path.to.save)
  
  # if (sheet == 1) {
  #   path.to.save <- paste0(path.to.save, "/arm") 
  # } else {
  #   path.to.save <- paste0(path.to.save, "/pol")
  # }
  
  path.to.save <- paste0(path.to.save, "/", lang)
  
  if (!dir.exists(path.to.save))
    dir.create(path.to.save)
  
  for (ptype in list(
    c("bar", 1050, 850), 
    c("density", 1000, 1200), 
    c("heat", 1250, 850)
  )) {
    png(
      paste0(path.to.save, "/", ptype[1], ".png"), 
      width = as.integer(ptype[2]), 
      height = as.integer(ptype[3]),
      res = 72
    )
    (
      plot(prct_matrix, ptype[1]) +
        theme(
          legend.title = element_blank(),
          axis.text = element_text(
            size = if (ptype[1] != "density") 12 else 14, 
            face = "bold", 
            color = "black"
          ),
          legend.text = element_text(size = 10, face = "bold"),
          strip.text = element_text(size = 12, face = "bold")
        ) + {
          if (ptype[1] == "bar") {
            theme(axis.title.x = element_blank())
          }
        }
    ) %>%
      print()
    dev.off()
  }
  
  path.to.save <- "./"
}
