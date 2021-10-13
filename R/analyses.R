library(car)
library(ggplot2)
library(gridExtra)
library(officer)
library(parallel)
library(readxl)
library(rvg)
library(writexl)

options(mc.cores = detectCores() - 1) 

# Working directory
setwd("~/Projects/Consultations/Favre Lucie (COOL-Saxenda)")

# Output directory
outdir <- paste0("results/analyses_", format(Sys.Date(), "%Y%m%d"))
if (!dir.exists(outdir)) dir.create(outdir)

# ------------------- Data importation and preprocessing -------------------- #

# Import data
fileName <- "data-raw/Saxenda  Database for Stat (1).xlsx"
sheets <- excel_sheets(fileName)
sheets <- setNames(sheets, tolower(sheets))
dta <- lapply(sheets, function(s) {
  na <- c("", "na", "NA", "à definir novembre")
  d <- suppressWarnings(read_xlsx(fileName, sheet = s, na = na))
  d <- as.data.frame(d)[apply(!is.na(d), 1, any), ]
  d[!is.na(d$ID), ]
})
rm(fileName, sheets)

# Help function - Determine time point from the data
tp <- function(d) {
  v <- regmatches(names(d), regexpr("M(0|4|10)$", names(d)))
  names(sort(table(v), decreasing = TRUE))[1]
}

# Recoding (1)
v <- "Dyslipidémie traité M4 (1=oui 0= non)" # A CONFIRMER
dta$m4[grepl("Elifem|Cerazette", dta$m4[[v]]) & !is.na(dta$m4[[v]]), v] <- 0
rm(v)

# Recoding (2)
dta <- lapply(dta, function(d) {
  for (v in names(d)[sapply(d, function(x) class(x)[1] == "character")]) {
    x <- trimws(d[[v]])
    if (all(grepl("^[0-9]+((\\.|,)[0-9]+)?$", x) | is.na(x))) {
      d[[v]] <- as.numeric(sub(",", ".", x))
    }
  }
  return(d)
})
if (FALSE) {
  str(dta$m0[sapply(dta$m0, function(x) class(x)[1] == "character")])
  str(dta$m4[sapply(dta$m4, function(x) class(x)[1] == "character")])
  str(dta$m10[sapply(dta$m10, function(x) class(x)[1] == "character")])
}

# Rename variables
dta <- lapply(dta, function(d) {
  V <- c("Phosphatase alcaline", "Prédiabète", "Lithiase vescilulaire",
         "Masse Totale")
  for (v in V) names(d) <- sub(v, paste(v, tp(d)), names(d))
  names(d) <- sub("^(Effets secondaires M(0|4|10))(.+)?", "\\1", names(d))
  names(d) <- sub("^(US abdo M(0|4|10))(.+)?", "\\1", names(d))
  names(d)[grepl("Consultation diet", names(d))] <- "Consultation diet n"
  names(d)[names(d) == "Fibroscan (kPa)"] <-
    paste("Fibroscan kPa", tp(d))
  names(d)[grepl("HbA1C M10 \\(mmol/L\\)", names(d))] <- "HbA1C M10 (%)"
  names(d) <- sub("^(.+)( (M4-M10|M(0|4|10)))(.+)", "\\1\\5\\2", names(d))
  names(d) <- sub(" M4-M10", ".M4.M10", names(d))
  names(d) <- sub(" (M(0|4|10))", ".\\1", names(d))
  names(d) <- sub("(\\()?%(\\))?", "pct", names(d))
  names(d) <- sub("\\s*\\([^\\)]+\\)", "", names(d))
  names(d) <- iconv(names(d), to = "ASCII//TRANSLIT")
  names(d) <- sub(" (mU/(l)?|pmol/l)", "", names(d))
  names(d) <- sub("FC/min", "FCmin", names(d))
  names(d) <- sub("SOAS IAH/h", "SOASiahh", names(d))
  names(d) <- sub(" 1=oui 0=non", "", names(d))
  names(d) <- gsub("( )(.)", "\\U\\2", names(d), perl = TRUE)
  d
})

# Age
dta$m0$Age <- as.numeric(as.Date(dta$m0$DateDebutTraitement) -
                           as.Date(dta$m0$DateDeNaissance)) / 365.2425

# Effets secondaires
dta[2:3] <- lapply(dta[2:3], function(d) {
  v <- grep("EffetsSecondaires", names(d), value = TRUE)
  m <- regmatches(v[1], regexpr("(4|10)$", v))
  if (m == "10") d[[v]] <- as.character(d[[v]])
  e <- lapply(lapply(strsplit(d[[v]], "(,|;)"), trimws), as.numeric)
  for (i in 1:10) {
    x <- paste0("EffetSecondaire", i, ".M", m)
    d[[x]] <- as.numeric(sapply(e, function(z) {
      if (any(is.na(z))) NA else i %in% z
    }))
  }
  return(d)
})

# Keep only numeric variable
lapply(dta, function(d) names(d)[!sapply(d, is.numeric)])
dta <- lapply(dta, function(d) {
  j1 <- sapply(d, function(x) class(x)[1] == "numeric")
  j2 <- !grepl("^Date", names(d))
  j3 <- !grepl("^EffetsSecondaires", names(d))
  d[j1 & j2 & j3]
})

# HTA M0
v <- c("ID", "HTA.M0", "HTATraite.M0")
tmp <- merge(dta$m0[v], dta$m4[v], by = "ID", all = TRUE)
if (all(with(tmp, HTA.M0.x == HTA.M0.y & HTATraite.M0.x == HTATraite.M0.y))) {
  dta$m4$HTA.M0 <- NULL
  dta$m4$HTATraite.M0 <- NULL
} else {
  stop("check HTA.M0")
}
rm(tmp, v)

# TSH
dta <- lapply(dta, function(d) {
  d[d$ID %in% c(949775, 999871, 3295571), grep("^TSH", names(d))] <- NA
  return(d)
})

# New variables
for (m in c("m4", "m10")) {
  M <- toupper(m)
  X <- "TSH"
  if (m == "m10") X <- c(X, "MasseTotale", "LeanMass", "LeanMassPct")
  dta[[m]] <- merge(dta[[m]], dta$m0[c("ID", paste0(X, ".M0"))], by = "ID",
                    all.x = TRUE)
  for (x in X) {
    dta[[m]][[paste0("Δ", x, ".", M)]] <-
      dta[[m]][[paste0(x, ".", M)]] - dta[[m]][[paste0(x, ".M0")]]
    dta[[m]][[paste0(x, ".M0")]] <- NULL
  }
}
rm(m, M, X, x)

# New variables - USAbdoBin
for (m in c("m0", "m10")) {
  x <- paste0("USAbdo.", toupper(m))
  y <- paste0("USAbdoBin.", toupper(m))
  dta[[m]][[y]] <- ifelse(dta[[m]][[x]] == 0, 0, 1)
}
rm(m, x, y)

# Rename variable (2) - Add missing M0, M4, M10
names(dta$m0)[names(dta$m0) == "SOAS"] <- "SOAS.M0"
names(dta$m0)[names(dta$m0) == "TourDeTaille"] <- "TourDeTaille.M0"
dta[2:3] <- lapply(dta[2:3], function(d) {
  m <- tp(d)
  b <- !grepl(paste0("\\.", m, "$"), names(d)) & names(d) != "ID"
  names(d)[b] <- paste0(names(d)[b], ".", m)
  return(d)
})

# -------------------------- Descriptive analyses --------------------------- #

# Analyses by time point
figs <- list()
tbls <- lapply(dta, function(d) {
  m <- tp(d)
  v_bin <- names(d)[sapply(d, function(x) all(is.na(x) | x %in% 0:1))]
  v_cat <- names(d)[sapply(d, function(x) {
    all(x %% 1 == 0 | is.na(x)) & length(unique(x)) <= 5
  })]
  v_cat <- v_cat[!(v_cat %in% v_bin)]
  v_num <- names(d)[!(names(d) %in% c("ID", v_bin, v_cat))]
  tbls <- list(
    bin = do.call(rbind, lapply(v_bin, function(v) {
      x <- d[[v]][!is.na(d[[v]])]
      data.frame(variable = v, nobs = length(x), npos = sum(x), prop = mean(x))
    })),
    cat = do.call(rbind, lapply(v_cat, function(v) {
      n <- table(d[[v]])
      p <- as.vector(prop.table(n))
      w <- names(n)
      n <- as.vector(n)
      data.frame(variable = v, value = w, n = n, prop = p)
    })),
    num = do.call(rbind, lapply(v_num, function(v) {
      x <- d[[v]][!is.na(d[[v]])]
      fig <- ggplot(data.frame(u = x), aes(x = factor(0), y = u)) +
        geom_boxplot() +
        scale_x_discrete(breaks = NULL) +
        labs(x = "", y = v, title = paste(m, v))
      figs <<- append(figs, setNames(list(fig), paste(m, v)))
      data.frame(
        variable = v,
        nobs = length(x),
        mean = mean(x),
        sd = sd(x),
        min = min(x),
        q25 = quantile(x, .25),
        median = median(x),
        q75 = quantile(x, .75),
        max = max(x, 75)
      )
    }))
  )
  lapply(tbls, function(tbl) {
    rownames(tbl) <- NULL
    return(tbl)
  })
})
write_xlsx(unlist(tbls, recursive = FALSE),
           file.path(outdir, "descriptive_analyses.xlsx"))

# Variables to compare
if (FALSE) {
  intersect(names(dta$m0), names(dta$m4))
  intersect(names(dta$m0), names(dta$m10))
  intersect(names(dta$m4), names(dta$m10))
}
Merge <- function(x1, x2) merge(x1, x2, by = "variable", all = TRUE)
cmp_var <- Reduce(Merge, mclapply(dta, function(d) {
  i <- grepl("M(0|4|10)$", names(d)) & !grepl("M4\\.M10$", names(d))
  v <- names(d)[i]
  m <- regmatches(v[1], regexpr("M(0|4|10)$", v[1]))
  v <- sub("\\.M(0|4|10)$", "", v)
  setNames(data.frame(v, 1), c("variable", m))
}))
cmp_var[is.na(cmp_var)] <- 0
cmp_var <- cmp_var[apply(cmp_var[2:4], 1, sum) > 1, ]
rm(Merge)

# Longitudinal data
lg <- Reduce(function(x1, x2) merge(x1, x2, by = "ID", all = TRUE), dta)

# Variable to compare - Type
cmp_var$type <- sapply(cmp_var$variable, function(v) {
  v <- paste0(v, ".M", c(0, 4, 10))
  v <- v[v %in% names(lg)]
  x <- do.call(c, lg[v])
  x <- x[!is.na(x)]
  if (all(x %in% 0:1)) {
    "bin"
  } else if (all(x %% 1 == 0) & length(unique(x)) <= 5) {
    "cat"
  } else {
    "num"
  }
})

# Comparisons
cmp_figs <- list()
M <- list(c("M0", "M4"), c("M0", "M10"), c("M4", "M10"))
M <- setNames(M, sapply(M, paste, collapse = "."))
cmp_tbls <- lapply(M, function(m) {
  types <- c("bin", "cat", "num")
  types <- setNames(types, types)
  lapply(types, function(type) {
    cv <- cmp_var[cmp_var$type == type & apply(cmp_var[m], 1, prod) == 1, ]
    if (nrow(cv) == 0) return(NULL)
    r <- do.call(rbind, lapply(1:nrow(cv), function(i) {
      v <- cv$variable[i]
      d <- na.omit(lg[paste(v, m, sep = ".")])
      if (type %in% c("bin", "cat")) {
        lvls <- unique(c(d[[1]], d[[2]]))
        x1 <- factor(d[[1]], lvls)
        x2 <- factor(d[[2]], lvls)
        pv <- if (length(lvls) > 1) mcnemar.test(x1, x2)$p.value else NA
        if (type == "bin") {
          r <- data.frame(
            variable = v,
            nobs = nrow(d),
            npos.1 = sum(d[[1]]),
            prop.1 = mean(d[[1]]),
            npos.2 = sum(d[[2]]),
            prop.2 = mean(d[[2]]),
            mcnemar.test.pv = pv
          )
        } else {
          n1 <- table(x1)
          p1 <- as.vector(prop.table(n1))
          n1 <- as.vector(n1)
          n2 <- table(x2)
          p2 <- as.vector(prop.table(n2))
          n2 <- as.vector(n2)
          r <- data.frame(
            variable = c(v, rep(NA, length(lvls) - 1)),
            value = lvls,
            n.1 = n1,
            prop.1 = p1,
            n.2 = n2,
            prop.2 = p2,
            mcnemar.test.pv = c(pv, rep(NA, length(lvls) - 1))
          )
        }
      } else {
        tmp <- rbind(data.frame(x = m[1], y = d[[1]]),
                     data.frame(x = m[2], y = d[[2]]))
        ttl <- paste(v, paste(m, collapse = " "))
        fig <- ggplot(tmp, aes(x = x, y = y)) +
          geom_boxplot() +
          labs(x = "", y = v, title = ttl)
        cmp_figs <<- append(cmp_figs, setNames(list(fig), ttl))
        if (nrow(d) > 1 ) {
          pv <- t.test(d[[1]], d[[2]], paired = TRUE)$p.value
        } else {
          pv <- NA
        }
        r <- data.frame(
          variable = v,
          n = nrow(d),
          mean.1 = mean(d[[1]]),
          sd.1 = sd(d[[1]]),
          mean.2 = mean(d[[2]]),
          sd.2 = sd(d[[2]]),
          mean_diff = mean(d[[2]] - d[[1]]),
          sd_diff = sd(d[[2]] - d[[1]]),
          paired.t.test.pv = pv
        )
      }
      return(r)
    }))
    names(r) <- sub("\\.1", paste0(".", m[1]), names(r))
    names(r) <- sub("\\.2", paste0(".", m[2]), names(r))
    return(r)
  })
})
tmp <- unlist(cmp_tbls, recursive = FALSE)
tmp <- tmp[!sapply(tmp, is.null)]
write_xlsx(tmp, file.path(outdir, "comparative_analyses.xlsx"))
rm(M, tmp)

# Export boxplots
if (TRUE) { # Takes some time
BP <- list(list(figs, "descriptive_analyses.pdf"),
           list(cmp_figs, "comparative_analyses.pdf"))
for (bp in BP) {
  f1 <- file.path(outdir, bp[[2]])
  cairo_pdf(f1, onefile = TRUE)
  for (fig in bp[[1]]) print(fig)
  dev.off()
  f2 <- "/tmp/tmp_bookmarks.txt"
  for(i in 1:length(bp[[1]])) {
    write("BookmarkBegin", file = f2, append = (i!=1))
    s <- paste("BookmarkTitle:", names(bp[[1]])[i])
    write(s, file = f2, append = TRUE)
    write("BookmarkLevel: 1", file = f2, append = TRUE)
    write(paste("BookmarkPageNumber:", i), file = f2, append = TRUE)
  }
  system(paste("pdftk", f1, "update_info", f2, "output /tmp/tmp_fig.pdf"))
  system(paste("mv /tmp/tmp_fig.pdf", f1, "&& rm", f2))
}
rm(BP, bp, fig, f1, f2, i, s)
}

# ---------------------------- Specific analyses ---------------------------- #

# Q1 : Prévalence perte de poids ≥5% et ≥10% à M4 et M10
Q1_prev_perte_poids <- matrix(ncol = 2, c(
  sum(!is.na(dta$m4$PerteDePoidsPct.M4)),
  mean(dta$m4$PerteDePoidsPct.M4 >= 5, na.rm = TRUE),
  mean(dta$m4$PerteDePoidsPct.M4 >= 10, na.rm = TRUE),
  sum(!is.na(dta$m10$PerteDePoidsPct.M10)),
  mean(dta$m10$PerteDePoidsPct.M10 >= 5, na.rm = TRUE),
  mean(dta$m10$PerteDePoidsPct.M10 >= 10, na.rm = TRUE)
))
Q1_prev_perte_poids <- cbind(
  data.frame(c("N", paste("Perte de poids >=", c("5%", "10%")))),
  Q1_prev_perte_poids
)
names(Q1_prev_perte_poids) <- c("", "M4", "M10")
write_xlsx(as.data.frame(Q1_prev_perte_poids),
           file.path(outdir, "Q1_prev_perte_poids.xlsx"))

# Q2 : Vérifier si existe une perte de poids (%) différente à M4 et M10 entre
#      hommes et femmes
V <- do.call(c, lapply(c(4, 10), function(m) {
  paste0("PerteDePoids", c("", "Pct"), ".M", m)
}))
lg$Sexe <- factor(lg$Sexe, 1:2, c("F", "H"))
Q2_perte_poids_Sexe <- do.call(rbind, lapply(V, function(v) {
  fml <- as.formula(paste(v, "~ Sexe"))
  n <- function(z, na.rm = FALSE) sum(!is.na(z))
  r <- do.call(c, lapply(c("n", "mean", "sd"), function(fct) {
    r <- aggregate(fml, lg, get(fct))
    setNames(r[[2]], paste(fct, r[[1]], sep = "."))
  }))[c(1, 3, 5, 2, 4, 6)]
  r <- c(r, t.test.p.value = t.test(fml, lg)$p.value,
         wilcox.test.p.value = wilcox.test(fml, lg, exact = FALSE)$p.value)
  cbind(data.frame(variable = v), t(r))
}))
write_xlsx(Q2_perte_poids_Sexe, file.path(outdir, "Q2_perte_poids_Sexe.xlsx"))

# Q3 : Comparer les différences de perte de poids M0-M4, M0-M10 (%) entre
#      BMI >35 et BMI<35 (BMI au temps 0)
lg$BMIcat.M0 <- factor(lg$BMI.M0 <= 35, c(TRUE, FALSE),
                       c("BMI.M0<=35", "BMI.M0>35"))
Q3_perte_poids_BMIcat <- do.call(rbind, lapply(V, function(v) {
  fml <- as.formula(paste(v, "~ BMIcat.M0"))
  n <- function(z, na.rm = FALSE) sum(!is.na(z))
  r <- do.call(c, lapply(c("n", "mean", "sd"), function(fct) {
    r <- aggregate(fml, lg, get(fct))
    setNames(r[[2]], paste(fct, r[[1]], sep = "."))
  }))[c(1, 3, 5, 2, 4, 6)]
  r <- c(r, t.test.p.value = t.test(fml, lg)$p.value,
         wilcox.test.p.value = wilcox.test(fml, lg, exact = FALSE)$p.value)
  cbind(data.frame(variable = v), t(r))
}))
write_xlsx(Q3_perte_poids_BMIcat,
           file.path(outdir, "Q3_perte_poids_BMIcat.xlsx"))
rm(V)

# Q6a : Relation entre delta TSH et perte de poids (%) (corrélation avec perte
#       de poids?)
Q6a_ΔTSH_PertePoids_figs <- list()
Q6a_ΔTSH_PertePoids <- do.call(rbind, lapply(1:3, function(k) {
  if (k %in% 1:2) {
    y <- "PerteDePoidsPct.M4"
    x <- "ΔTSH.M4"
  } else {
    y <- "PerteDePoidsPct.M10"
    x <- "ΔTSH.M10"
  }
  if (k == 2) {
    d <- na.omit(lg[!(lg$ID == 3295571), c(x, y)])
    com <- "Without ID 3295571"
  } else {
    d <- na.omit(lg[c(x, y)])
    com <- ""
  }
  delta <- qnorm(0.975) / sqrt(nrow(d) - 3)
  cor_pearson <- cor(d[[x]], d[[y]])
  cor_pearson_lwr <- tanh(atanh(cor_pearson) - delta)
  cor_pearson_upr <- tanh(atanh(cor_pearson) + delta)
  cor_pearson_pv <- cor.test(d[[x]], d[[y]])$p.value
  cor_spearman <- cor(d[[x]], d[[y]], method = "spearman")
  cor_spearman_lwr <- tanh(atanh(cor_spearman) - delta)
  cor_spearman_upr <- tanh(atanh(cor_spearman) + delta)
  cor_spearman_pv <- cor.test(d[[x]], d[[y]], method = "spearman",
                              exact = FALSE)$p.value
  fit <- lm(as.formula(paste(y, "~", x)), d)
  b <- signif(coef(fit)[[2]], 3)
  ci <- signif(confint(fit)[2, ], 3)
  p <- coef(summary(fit))[2, "Pr(>|t|)" ]
  p <- if (p >= 0.001) paste0("p=", round(p, 3)) else "p<0.001"
  r2 <- paste0("R2=", round(summary(fit)$r.squared, 3))
  cap <- paste0("slope = ", b, " (", ci[1], ",", ci[2], "), ", p, ", ", r2)
  fig <- ggplot(d, aes_string(y = y, x = x)) +
    geom_point() +
    geom_smooth(method = lm, formula = y ~ x, se = FALSE) +
    labs(caption = cap)# +
    #theme(axis.title = element_text(size = rel(.75)))
  if (k == 2) fig <- fig + labs(subtitle = com)
  Q6a_ΔTSH_PertePoids_figs <<- append(Q6a_ΔTSH_PertePoids_figs, list(fig))
  data.frame(
    variable1 = x,
    variable2 = y,
    comment = com,
    nobs = nrow(d),
    cor_pearson = cor_pearson,
    cor_pearson_lwr = cor_pearson_lwr,
    cor_pearson_upr = cor_pearson_upr,
    cor_pearson_pv = cor_pearson_pv,
    cor_spearman = cor_spearman,
    cor_spearman_lwr = cor_spearman_lwr,
    cor_spearman_upr = cor_spearman_upr,
    cor_spearman_pv = cor_spearman_pv
  )
}))
write_xlsx(Q6a_ΔTSH_PertePoids,
           file.path(outdir, "Q6a_ΔTSH_PertePoids.xlsx"))
cairo_pdf(file.path(outdir, "Q6a_ΔTSH_PertePoids.pdf"), onefile = TRUE)
for (fig in Q6a_ΔTSH_PertePoids_figs) print(fig)
dev.off()
rm(fig)

# Q7 Prédicteurs de la perte de poids (%) (variable dépendante) à M4 et à M10 ?
#    (variables : BMI M0, FMI M0, FM% M0, VAT M0, age, sexe, BITE)
Y <- c("PerteDePoidsPct.M4", "PerteDePoidsPct.M10")
X <- c("BMI.M0", "FMI.M0", "FatMassPct.M0", "VAT.M0", "Age", "Sexe", "BITE.M0")
Q7_uni_reg_tbl <- do.call(rbind, lapply(X, function(x) {
  do.call(rbind, lapply(Y, function(y) {
    d <- na.omit(lg[c(x, y)])
    fit <- lm(as.formula(paste(y, "~", x)), d)
    r <- c(coef(fit), as.vector(t(confint(fit))))[c(1, 3:4, 2, 5:6)]
    names(r) <- c(paste0("Intercept", c("", ".lwr", ".upr")),
                  paste0("Slope", c("", ".lwr", ".upr")))
    r <- c(r, p.value = anova(fit)$`Pr(>F)`[1])
    s <- data.frame(
      dependent_variable = y,
      independent_variable = x,
      nobs = nrow(fit$model)
    )
    cbind(s, t(r))
  }))
}))
write_xlsx(Q7_uni_reg_tbl,
           file.path(outdir, "Q7_univariable_regressions.xlsx"))

# Q7 - Figures of the univariable regressions
Q7_uni_reg_figs <- mclapply(setNames(X, X), function(x) {
  lapply(setNames(Y, Y), function(y) {
    d <- na.omit(lg[c(x, y)])
    fit <- lm(as.formula(paste(y, "~", x)), d)
    b <- signif(coef(fit)[[2]], 3)
    ci <- signif(confint(fit)[2, ], 3)
    p <- coef(summary(fit))[2, "Pr(>|t|)" ]
    p <- if (p >= 0.001) paste0("p=", round(p, 3)) else "p<0.001"
    r2 <- paste0("R2=", round(summary(fit)$r.squared, 3))
    cap <- paste0("b = ", b, " (", ci[1], ",", ci[2], "), ", p, ", ", r2)
    fig <- ggplot(d, aes_string(y = y, x = x))
    if (x == "Sexe") {
      fig <- fig + geom_boxplot()
    } else {
      fig <- fig +
        geom_point() +
        geom_smooth(method = lm, formula = y ~ x, se = FALSE)
    }
    fig + labs(caption = cap)
  })
})
pdf(file.path(outdir, "Q7_univariable_regressions.pdf"), width = 12)
for (figs in Q7_uni_reg_figs) {
  do.call(grid.arrange, append(figs, list(ncol = 2)))
}
dev.off()
rm(figs)

# Q7 - Multivariate regressions
if (FALSE) {
  require(GGally)
  ggpairs(lg[X])
}
Q7_multi_reg <- mclapply(setNames(Y, Y), function(y) {
  X <- c("VAT.M0", "Age", "Sexe", "BITE.M0")
  W <- c("BMI.M0", "FMI.M0", "FatMassPct.M0")
  lapply(setNames(W, W), function(w) {
    X <- c(X, w)
    fml <- as.formula(paste(y, "~", paste(X, collapse = " + ")))
    longitudinal_data <- na.omit(lg[c(y, X)])
    fit <- do.call("lm", list(formula = fml, data = quote(longitudinal_data)))
    tbl <- cbind(data.frame(coefficient = names(coef(fit)), beta = coef(fit)),
                 confint(fit), `p-value` = coef(summary(fit))[, 4])
    tbl0 <- do.call(rbind, lapply(names(fit$model)[-1], function(x) {
      fit0 <- lm(paste(y, "~", x), fit$model)
      cbind(data.frame(coefficient = names(coef(fit0)), beta = coef(fit0)),
            confint(fit0), `p-value` = coef(summary(fit0))[, 4])[-1, ]
    }))
    names(tbl0)[-1] <- paste(names(tbl0)[-1], "(univar)")
    vif <- vif(fit)
    vif <- data.frame(coefficient = sub("Sexe", "SexeH", names(vif)),
                      vif = vif)
    tbl$dummy_row_number <- 1:nrow(tbl)
    tbl <- merge(tbl0, tbl, by = "coefficient", all = TRUE)
    tbl <- merge(tbl, vif, by = "coefficient", all = TRUE)
    tbl <- tbl[order(tbl$dummy_row_number), ]
    tbl$dummy_row_number <- NULL
    tbl <- cbind(dependant_variable = c(y, rep(NA, nrow(tbl) - 1)),
                 nobs = c(nrow(fit$model), rep(NA, nrow(tbl) - 1)),
                 tbl)
    list(fit = fit, tbl = tbl, n = nrow(fit$model))
  })
})
tmp <- lapply(unlist(Q7_multi_reg, recursive = FALSE), function(z) z$tbl)
names(tmp) <- sub("PerteDePoidsPct\\.", "", names(tmp))
write_xlsx(tmp, file.path(outdir, "Q7_multivariable_regressions.xlsx"))
get_fml <- function(fit) {
  paste(as.character(fit$call$formula)[c(2, 1, 3)], collapse = " ")
}
tmp <- lapply(unlist(Q7_multi_reg, recursive = FALSE), function(z) z$fit)
pdf(file.path(outdir, "Q7_multivariable_regressions.pdf"))
for (fit in tmp) {
  par(mfrow = c(2, 2))
  for (i in 1:4) plot(fit, i)
  par(mfrow = c(1, 1))
  mtext(get_fml(fit), outer = TRUE, line = -1.8, cex = 1)
}
dev.off()
rm(X, Y, tmp, get_fml, fit, i)

# Q8 : Prédicteurs du pourcentage de perte masse maigre sur kg (variable
#      dépendante) à M10 ? (Variables : BMI M0, lean mass M0, masse totale
#      perdue M10, age, sexe)
y <- "LeanMassPct.M10"
X <- c("BMI.M0", "LeanMass.M0", "ΔMasseTotale.M10", "Age", "Sexe")
Q8_uni_reg_tbl <- do.call(rbind, lapply(X, function(x) {
  d <- na.omit(lg[c(x, y)])
  fit <- lm(as.formula(paste(y, "~", x)), d)
  r <- c(coef(fit), as.vector(t(confint(fit))))[c(1, 3:4, 2, 5:6)]
  names(r) <- c(paste0("Intercept", c("", ".lwr", ".upr")),
                paste0("Slope", c("", ".lwr", ".upr")))
  r <- c(r, p.value = anova(fit)$`Pr(>F)`[1])
  s <- data.frame(
    dependent_variable = y,
    independent_variable = x,
    nobs = nrow(fit$model)
  )
  cbind(s, t(r))
}))
write_xlsx(Q8_uni_reg_tbl,
           file.path(outdir, "Q8_univariable_regressions.xlsx"))

# Q8 - Figures of the univariable regressions
Q8_uni_reg_figs <- mclapply(setNames(X, X), function(x) {
  d <- na.omit(lg[c(x, y)])
  fit <- lm(as.formula(paste(y, "~", x)), d)
  b <- signif(coef(fit)[[2]], 3)
  ci <- signif(confint(fit)[2, ], 3)
  p <- coef(summary(fit))[2, "Pr(>|t|)" ]
  p <- if (p >= 0.001) paste0("p=", round(p, 3)) else "p<0.001"
  r2 <- paste0("R2=", round(summary(fit)$r.squared, 3))
  cap <- paste0("b = ", b, " (", ci[1], ",", ci[2], "), ", p, ", ", r2)
  fig <- ggplot(d, aes_string(y = y, x = x))
  if (x == "Sexe") {
    fig <- fig + geom_boxplot()
  } else {
    fig <- fig +
      geom_point() +
      geom_smooth(method = lm, formula = y ~ x, se = FALSE)
  }
  fig + labs(caption = cap)
})
cairo_pdf(file.path(outdir, "Q8_univariable_regressions.pdf"), onefile = TRUE)
for (fig in Q8_uni_reg_figs) print(fig)
dev.off()
rm(fig)

# Q8 - Multivariate regressions
if (FALSE) {
  require(GGally)
  ggpairs(lg[X])
}
Q8_multi_reg <- mclapply(1:3, function(k) {
  if (k == 2) {
    X <- X[X != "ΔMasseTotale.M10"]
  } else if (k == 3) {
    X <- X[X != "BMI.M0"]
  }
  fml <- as.formula(paste(y, "~", paste(X, collapse = " + ")))
  longitudinal_data <- na.omit(lg[c(y, X)])
  fit <- do.call("lm", list(formula = fml, data = quote(longitudinal_data)))
  tbl <- cbind(data.frame(coefficient = names(coef(fit)), beta = coef(fit)),
               confint(fit), `p-value` = coef(summary(fit))[, 4])
  tbl0 <- do.call(rbind, lapply(names(fit$model)[-1], function(x) {
    fit0 <- lm(paste(y, "~", x), fit$model)
    cbind(data.frame(coefficient = names(coef(fit0)), beta = coef(fit0)),
          confint(fit0), `p-value` = coef(summary(fit0))[, 4])[-1, ]
  }))
  names(tbl0)[-1] <- paste(names(tbl0)[-1], "(univar)")
  vif <- vif(fit)
  vif <- data.frame(coefficient = sub("Sexe", "SexeH", names(vif)),
                    vif = vif)
  tbl$dummy_row_number <- 1:nrow(tbl)
  tbl <- merge(tbl0, tbl, by = "coefficient", all = TRUE)
  tbl <- merge(tbl, vif, by = "coefficient", all = TRUE)
  tbl <- tbl[order(tbl$dummy_row_number), ]
  tbl$dummy_row_number <- NULL
  tbl <- cbind(dependant_variable = c(y, rep(NA, nrow(tbl) - 1)),
               nobs = c(nrow(fit$model), rep(NA, nrow(tbl) - 1)),
               tbl)
  list(fit = fit, tbl = tbl, n = nrow(fit$model))
})
names(Q8_multi_reg) <- paste0("Model", 1:3)
write_xlsx(lapply(Q8_multi_reg, function(z) z$tbl),
           file.path(outdir, "Q8_multivariable_regressions.xlsx"))
cairo_pdf(file.path(outdir, "Q8_multivariable_regressions.pdf"),
          onefile = TRUE)
for (s in names(Q8_multi_reg)) {
  par(mfrow = c(2, 2))
  for (i in 1:4) plot(Q8_multi_reg[[s]]$fit, i)
  par(mfrow = c(1, 1))
  mtext(s, outer = TRUE, line = -1.8, cex = 1)
}
dev.off()
rm(X, y, s, i)

# Q10: Graphique VAT, lean mass (g), fat mass (g) à M0 et M10
Y <- c("VAT", "FatMassPct", "LeanMassPct", "FatMass", "LeanMass")
Q10_figs <- mclapply(setNames(Y, Y), function(y) {
  tmp <- subset(cmp_tbls$M0.M10$num, variable == y)
  tmp$se.M0 <- tmp$sd.M0 / sqrt(tmp$n)
  tmp$se.M10 <- tmp$sd.M10 / sqrt(tmp$n)
  tmp <- tmp[grep("^(n|(mean|sd|se)\\.M1?0)$", names(tmp))]
  m <- c("M0", "M10")
  s <- c("mean", "sd", "se")
  tmp <- reshape(tmp, varying = lapply(s, paste, m, sep = "."), v.names = s,
                 times = m, direction = "long")
  tmp$id <- NULL
  tmp <- cbind(variable = c(y, NA), tmp)
  yLab <- c(VAT = "VAT (g)", FatMassPct = "Fat Mass (%)",
            LeanMassPct = "Lean Mass (%)", FatMass = "Fat Mass (g)",
            LeanMass = "Lean Mass (g)")
  fig <- ggplot(tmp, aes(x = time, y = mean, fill = time)) +
    geom_bar(position = position_dodge(width = 0.9), stat = "identity", 
             color="black") +
    geom_errorbar(aes(ymax = mean + se, ymin = mean), width = 0.5,
                  position = position_dodge(width = 0.9)) +
    theme_classic() +
    scale_fill_manual(values = c("white", "black")) +
    guides(fill = "none") +
    labs(x = "", y = yLab[y])
  attr(fig, "data") <- tmp
  return(fig)
})
for (y in names(Q10_figs)) {
  tiff(file.path(outdir, paste0("Q10_fig_", y, ".tiff")),
       width = 2400, height = 4800, res = 1152, compression = "zip")
  print(Q10_figs[[y]])
  dev.off()
  doc <- read_pptx()
  doc <- add_slide(doc, 'Title and Content', 'Office Theme')
  anyplot <- dml(ggobj = Q10_figs[[y]])
  doc <- ph_with(doc, anyplot, location = ph_location_fullsize())
  print(doc, target = file.path(outdir, paste0("Q10_fig_", y, ".pptx")))
  write_xlsx(attr(Q10_figs[[y]], "data"),
             file.path(outdir, paste0("Q10_fig_", y, ".xlsx")))
}
rm(y, doc, anyplot)

# Q11 : Comparaison pourcentage perte masse maigre sut kg total entre hommes
#       versus femmes
v <- "PourcentagePerteMasseMaigreSutKgTotal.M10"
Y <- lg[[v]]
Y <- list(All = Y, Women = Y[lg$Sexe == "F"], Men = Y[lg$Sexe == "H"])
Y <- lapply(Y, na.omit)
n <- function(x) length(x)
q25 <- function(x) quantile(x, .25)[[1]]
q75 <- function(x) quantile(x, .75)[[1]]
fcts <- c("n", "mean", "sd", "min", "q25", "median", "q75", "max")
Q11_table <- t(sapply(Y, function(y) sapply(fcts, function(fct) get(fct)(y))))
fml <- PourcentagePerteMasseMaigreSutKgTotal.M10 ~ Sexe
pv1 <- t.test(fml, lg)$p.value
pv2 <- wilcox.test(fml, lg, exact = FALSE)$p.value
Q11_table <- cbind(Q11_table, t.test.pval = c(pv1, NA, NA),
                   wilcox.test.pval = c(pv2, NA, NA))
Q11_table <- cbind(data.frame(Variable = c(v, NA, NA),
                              Group = rownames(Q11_table)),
                   Q11_table)
write_xlsx(Q11_table, file.path(outdir, "Q11_table.xlsx"))
pdf(file.path(outdir, "Q11_boxplot.pdf"))
ggplot(na.omit(lg[c(v, "Sexe")]), aes_string(x = "Sexe", y = v)) +
  geom_boxplot()
dev.off()

# Q12 : Analyses de régression multivariable
#       leanmasspct: sexe, BMI, age
#       pertedepoidspct: FatMassPct.M0, sexe, age
#       fattmasspct: sexe, BMI, age
Y <- c("PerteDePoidsPct.M4", "PerteDePoidsPct.M10", "LeanMassPct.M10",
       "FatMassPct.M10")
Q12_multi_reg <- mclapply(setNames(Y, Y), function(y) {
  X <- if (grepl("^Perte", y)) "FatMassPct.M10" else "BMI.M0"
  X <- c(X, "Age", "Sexe")
  fml <- as.formula(paste(y, "~", paste(X, collapse = " + ")))
  longitudinal_data <- na.omit(lg[c(y, X)])
  fit <- do.call("lm", list(formula = fml, data = quote(longitudinal_data)))
  tbl <- cbind(data.frame(coefficient = names(coef(fit)), beta = coef(fit)),
               confint(fit), `p-value` = coef(summary(fit))[, 4])
  tbl0 <- do.call(rbind, lapply(names(fit$model)[-1], function(x) {
    fit0 <- lm(paste(y, "~", x), fit$model)
    cbind(data.frame(coefficient = names(coef(fit0)), beta = coef(fit0)),
          confint(fit0), `p-value` = coef(summary(fit0))[, 4])[-1, ]
  }))
  names(tbl0)[-1] <- paste(names(tbl0)[-1], "(univar)")
  vif <- vif(fit)
  vif <- data.frame(coefficient = sub("Sexe", "SexeH", names(vif)),
                    vif = vif)
  tbl$dummy_row_number <- 1:nrow(tbl)
  tbl <- merge(tbl0, tbl, by = "coefficient", all = TRUE)
  tbl <- merge(tbl, vif, by = "coefficient", all = TRUE)
  tbl <- tbl[order(tbl$dummy_row_number), ]
  tbl$dummy_row_number <- NULL
  tbl <- cbind(dependant_variable = c(y, rep(NA, nrow(tbl) - 1)),
               nobs = c(nrow(fit$model), rep(NA, nrow(tbl) - 1)),
               tbl)
  list(fit = fit, tbl = tbl, n = nrow(fit$model))
})
write_xlsx(lapply(Q12_multi_reg, function(z) z$tbl),
           file.path(outdir, "Q12_multivariable_regressions.xlsx"))
cairo_pdf(file.path(outdir, "Q12_multivariable_regressions.pdf"),
          onefile = TRUE)
for (s in names(Q12_multi_reg)) {
  par(mfrow = c(2, 2))
  for (i in 1:4) plot(Q12_multi_reg[[s]]$fit, i)
  par(mfrow = c(1, 1))
  mtext(s, outer = TRUE, line = -1.8, cex = 1)
}
dev.off()
rm(Y, s, i)

# --------------------------------------------------------------------------- #

# Session Info
sink(file.path(outdir, "sessionInfo.txt"))
print(sessionInfo(), locale = TRUE)
sink()

