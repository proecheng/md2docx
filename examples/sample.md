# 公式示例文档 / Formula Examples

本文档包含各种数学公式，用于测试 MD2DOCX 转换器。

This document contains various mathematical formulas to test the MD2DOCX converter.

## 块级公式 / Block Formulas

### 分数与求和 / Fractions and Summations

$$\frac{a + b}{c - d}$$

$$\sum_{i=1}^{n} x_i = x_1 + x_2 + \cdots + x_n$$

### 积分 / Integrals

$$\int_0^{\infty} e^{-x^2} dx = \frac{\sqrt{\pi}}{2}$$

### 矩阵 / Matrices

$$A = \begin{pmatrix} a & b \\ c & d \end{pmatrix}$$

### 复杂公式 / Complex Formulas

$$L(\theta) = \mathbb{E}_{(s,a) \sim \pi_{\theta_{old}}} \left[ \frac{\pi_\theta(a|s)}{\pi_{\theta_{old}}(a|s)} A^{\pi_{\theta_{old}}}(s,a) \right]$$

## 行内公式 / Inline Formulas

著名的质能方程 $E = mc^2$ 由爱因斯坦提出。

The quadratic formula is $x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$.

函数 $f(x) = \sin(x) + \cos(x)$ 是周期函数。

## 希腊字母 / Greek Letters

- α (alpha) - 学习率
- β (beta) - 折扣因子
- γ (gamma) - 衰减率
- θ (theta) - 参数
- π (pi) - 策略函数
- λ (lambda) - 正则化系数

## 下标与上标 / Subscripts and Superscripts

变量 $x_t$ 表示时刻 $t$ 的状态。

$W^{(k)}$ 是第 $k$ 层的权重矩阵。

$h_i^{(l)}$ 表示第 $l$ 层节点 $i$ 的隐藏状态。

## 表格 / Tables

| 符号 | 含义 | 公式 |
|------|------|------|
| $\alpha$ | 学习率 | $\theta \leftarrow \theta - \alpha \nabla L$ |
| $\gamma$ | 折扣因子 | $R = \sum_{t=0}^{\infty} \gamma^t r_t$ |
| $\epsilon$ | 探索率 | $\epsilon$-greedy 策略 |

---

转换完成后，请在 Word 中双击任意公式以验证其可编辑性。

After conversion, double-click any formula in Word to verify it's editable.
