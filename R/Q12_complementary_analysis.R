# Model
fit <- lm(PerteDePoidsPct.M4 ~ FatMassPct.M0 + Sexe, lg)
summary(fit)

# Predictions
ndta <- expand.grid(Sexe = 1:2, FatMassPct.M0 = seq(33, 62, .1))
ndta$PerteDePoidsPct.M4 <- predict(fit, newdata = ndta)

# FatMassPct.M0 ~ Sexe
t.test(FatMassPct.M0 ~ Sexe, data = fit$model)
pdf("~/Boxplots_FatMassPct.M0_Sexe.pdf")
ggplot(fit$model, aes(x = factor(Sexe), y = FatMassPct.M0)) +
  geom_boxplot() +
  labs(x = "Sexe")
dev.off()

# Figure PerteDePoidsPct.M4 ~ FatMassPct.M0 + Sexe
png("~/ScatterPlot_PerteDePoidsPct.M4_FatMassPct.M0_Sexe.png")
ggplot(fit$model, aes(x = FatMassPct.M0, y = PerteDePoidsPct.M4)) +
  geom_point() +
  geom_smooth(method = "lm", formula = y ~ x, se = FALSE) +
  geom_line(data = ndta) +
  facet_wrap(~Sexe)
dev.off()

# Model with interaction
fit <- lm(PerteDePoidsPct.M4 ~ FatMassPct.M0 * Sexe,
          within(lg, {Sexe = factor(Sexe)}))
summary(fit)

