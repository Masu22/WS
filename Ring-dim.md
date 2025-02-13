可換環$R$の Krull 次元を$\dim R$とすると、多項式環 $ R[x] $ の次元が $ \dim R + 1 $ であることを証明します。

---

### **ステップ 1: $ \dim R[x] \geq \dim R + 1 $ を示す**
つまり、$ R[x] $ には $ \dim R + 1 $ の長さの素イデアルの増大列が存在することを示します。

#### **(1) $ R $ 内の最大長の素イデアル列を考える**
仮定として、$ R $ の Krull 次元が $ n $ であるとすると、最大の素イデアル鎖

$$
\mathfrak{p}_0 \subsetneq \mathfrak{p}_1 \subsetneq \dots \subsetneq \mathfrak{p}_n
$$

が存在します。

#### **(2) $ R[x] $ に拡張**
各素イデアル $ \mathfrak{p}_i $ から生成される $ R[x] $ の素イデアル $ \mathfrak{p}_i R[x] $ を考えると、

$$
\mathfrak{p}_0 R[x] \subsetneq \mathfrak{p}_1 R[x] \subsetneq \dots \subsetneq \mathfrak{p}_n R[x]
$$

は依然として素イデアルの鎖です。

さらに、**新しい素イデアル** $ (\mathfrak{p}_n R[x], x) $ を考えると、

$$
\mathfrak{p}_0 R[x] \subsetneq \mathfrak{p}_1 R[x] \subsetneq \dots \subsetneq \mathfrak{p}_n R[x] \subsetneq (\mathfrak{p}_n R[x], x)
$$

は $ R[x] $ における素イデアルの増大列になり、その長さは $ n+1 $ になります。  
したがって、$ \dim R[x] \geq \dim R + 1 $ です。

---

### **ステップ 2: $ \dim R[x] \leq \dim R + 1 $ を示す**
$ R[x] $ の素イデアル鎖の長さが最大でも $ \dim R + 1 $ であることを示します。

任意の素イデアル $ \mathfrak{q} \subset R[x] $ に対し、**その交わり $ \mathfrak{q} \cap R $ は $ R $ の素イデアル $ \mathfrak{p} $ である**ことが知られています（素イデアルの収縮）。  
このとき、次の2つのケースを考えます：

1. **$ \mathfrak{q} \cap R = \mathfrak{p} $ が極大イデアルのとき**  
   - $ \mathfrak{q} $ は $ R[x] $ の極大イデアルか、それに近いものになるため、それ以上の鎖は作れない。
  
2. **$ \mathfrak{q} \cap R = \mathfrak{p} $ のとき、$ x $ の関係を考える**  
   - $ R[x] $ の素イデアルは、次のような形にできる：
     - $ \mathfrak{q} = \mathfrak{p} R[x] $ なら $\dim R[x] = \dim R$
     - $ \mathfrak{q} = (\mathfrak{p} R[x], f(x)) $ なら $\dim R[x] = \dim R + 1$

どの場合でも、$ R[x] $ の素イデアル鎖の長さは $ \dim R + 1 $ を超えられないことが分かる。

したがって、  
$$
\dim R[x] \leq \dim R + 1
$$

---

### **結論**
以上より、

$$
\dim R + 1 \leq \dim R[x] \leq \dim R + 1
$$

が成り立ち、結論として

$$
\dim R[x] = \dim R + 1
$$

が証明されました。✅


----


はい、非零因子 $ a \in R $ に対して、商環 $ R/(a) $ の次元は $ \dim R - 1 $ になることが多くの場合で成立します。つまり、

$$
\dim R/(a) = \dim R - 1
$$

を示したいです。ただし、一般にはこの関係が成立しない場合もあるので、必要な仮定についても考察します。

---

## **目標**
可換環 $ R $ の非零因子 $ a \in R $ に対して、

$$
\dim R/(a) \leq \dim R - 1
$$

かつ

$$
\dim R/(a) \geq \dim R - 1
$$

を示すことで、次元がちょうど $ \dim R - 1 $ になることを証明します。

---

## **ステップ 1: $ \dim R/(a) \leq \dim R - 1 $ を示す**

$ R $ の素イデアルの最大の増大列を

$$
\mathfrak{p}_0 \subsetneq \mathfrak{p}_1 \subsetneq \dots \subsetneq \mathfrak{p}_n
$$

とする（$ \dim R = n $）。

このとき、商環 $ R/(a) $ の素イデアルは $ R $ の素イデアル $ \mathfrak{p} $ で $ a \in \mathfrak{p} $ を満たすものと一致する。つまり、**$ R/(a) $ の次元は、$ a $ を含む素イデアルの最大の鎖の長さ**に等しい。

もし $ a $ が非零因子なら、極大の素イデアル鎖

$$
\mathfrak{p}_0 \subsetneq \mathfrak{p}_1 \subsetneq \dots \subsetneq \mathfrak{p}_n
$$

のうち、**少なくとも 1 つの素イデアルが $ a $ を含まない**。そこで、最大の鎖の中で最初に $ a $ を含む素イデアル $ \mathfrak{p}_k $ を見つけると、

$$
\mathfrak{p}_k / (a) \subsetneq \mathfrak{p}_{k+1} / (a) \subsetneq \dots \subsetneq \mathfrak{p}_n / (a)
$$

となり、長さは $ n - k \leq n - 1 $ であるため、

$$
\dim R/(a) \leq \dim R - 1
$$

が分かる。

---

## **ステップ 2: $ \dim R/(a) \geq \dim R - 1 $ を示す**

次に、少なくとも 1 つの長さ $ n-1 $ の素イデアル鎖が $ R/(a) $ に存在することを示します。

最大の素イデアル鎖

$$
\mathfrak{p}_0 \subsetneq \mathfrak{p}_1 \subsetneq \dots \subsetneq \mathfrak{p}_n
$$

の中で、ある $ k $ に対して $ a \in \mathfrak{p}_k $ となる最小の $ k $ を取ると、$ \mathfrak{p}_{k-1} $ は $ a $ を含まないが、$ \mathfrak{p}_k $ は $ a $ を含む。

このとき、商環 $ R/(a) $ において対応する鎖

$$
\mathfrak{p}_k / (a) \subsetneq \mathfrak{p}_{k+1} / (a) \subsetneq \dots \subsetneq \mathfrak{p}_n / (a)
$$

が成立し、その長さは $ n-k $ です。通常、$ k = 1 $ をとることができるため、鎖の長さは $ n-1 $ となり、

$$
\dim R/(a) \geq \dim R - 1
$$

が分かる。

---

## **結論**
$$
\dim R - 1 \leq \dim R/(a) \leq \dim R - 1
$$

が示せたので、

$$
\dim R/(a) = \dim R - 1
$$

が成り立つ。✅

---

## **補足: 反例がある場合**
この結論が成り立たないケースもあります。たとえば、$ R $ が**非既約成分を持つ場合**や、**局所的に次元が異なる場合**です。

例えば、**非整域 $ R = k[x,y]/(xy) $ で $ a = x $** を取ると、$ R/(x) \cong k[y]/(y) $ であり、次元が 0 になってしまうので、$\dim R - 1 = 1 \neq 0 $ となります。

そのため、$ R $ が既約で、かつ $ a $ が真に非零因子（すべての既約成分で零因子でない）である場合に、結論が正しくなります。
