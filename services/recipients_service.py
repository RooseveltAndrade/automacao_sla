import pandas as pd
import unicodedata


def _norm_txt(v) -> str:
    s = str(v or "").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.upper()


def _norm_col(name: str) -> str:
    return " ".join(_norm_txt(name).split())


class RecipientsService:
    """
    Lê a planilha data/lideres.xlsx e devolve os e-mails por regional.
    - tenta ser resiliente com nomes de colunas
    """

    def __init__(self, xlsx_path: str, sheet_name: str | None = None):
        self.xlsx_path = xlsx_path
        self.df = pd.read_excel(self.xlsx_path, sheet_name=sheet_name)

        # Normaliza nomes de colunas para busca flexível
        self.col_map = {_norm_col(c): c for c in self.df.columns}

    def _find_col(self, *candidates: str) -> str | None:
        for c in candidates:
            key = _norm_col(c)
            if key in self.col_map:
                return self.col_map[key]
        return None

    def get_emails_by_regional(self, regional: str) -> list[str]:
        # tenta achar coluna de regional
        col_reg = self._find_col("REGIONAL", "INTEGRADA", "REGIAO", "REGIÃO", "UF", "NOME_REGIONAL")
        if not col_reg:
            raise RuntimeError("Não encontrei coluna de REGIONAL na planilha. Me diga o nome exato da coluna.")

        # tenta achar colunas de email
        col_email_ger = self._find_col("EMAIL_GERENTE", "GERENTE_EMAIL", "EMAIL GERENTE", "E-MAIL GERENTE")
        col_email_dir = self._find_col("EMAIL_DIRETOR", "DIRETOR_EMAIL", "EMAIL DIRETOR", "E-MAIL DIRETOR")
        col_email_apo1 = self._find_col("EMAIL_APOIO_1", "EMAIL APOIO 1", "E-MAIL APOIO 1")
        col_email_apo2 = self._find_col("EMAIL_APOIO_2", "EMAIL APOIO 2", "E-MAIL APOIO 2")
        col_email_apo = self._find_col("EMAIL_APOIO", "APOIO_EMAIL", "EMAIL APOIO", "E-MAIL APOIO")
        col_email_unico = self._find_col("EMAIL", "E-MAIL", "MAIL")

        df = self.df.copy()
        df[col_reg] = df[col_reg].astype(str).str.strip().str.upper()
        reg = str(regional).strip().upper()

        row = df[df[col_reg] == reg]
        if row.empty:
            return []

        row = row.iloc[0]
        emails = []

        def add_email(val):
            if val is None:
                return
            s = str(val).strip()
            s_upper = s.upper()
            if not s or s.lower() == "nan":
                return
            if s_upper.startswith("SEM_"):
                return
            # Permite mais de um email na mesma celula (separado por virgula ou ponto-e-virgula)
            for part in s.replace(";", ",").split(","):
                email = part.strip()
                if not email:
                    continue
                if email.upper().startswith("SEM_"):
                    continue
                emails.append(email)

        # se existir coluna única, usa ela
        if col_email_unico:
            add_email(row.get(col_email_unico))
        else:
            if col_email_ger:
                add_email(row.get(col_email_ger))
            if col_email_dir:
                add_email(row.get(col_email_dir))
            if col_email_apo1:
                add_email(row.get(col_email_apo1))
            if col_email_apo2:
                add_email(row.get(col_email_apo2))
            if col_email_apo:
                add_email(row.get(col_email_apo))

        # remove duplicados mantendo ordem
        seen = set()
        out = []
        for e in emails:
            if e not in seen:
                seen.add(e)
                out.append(e)
        return out