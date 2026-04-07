import pandas as pd
import unicodedata


def _norm_txt(v) -> str:
    s = str(v or "").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.upper()


def _norm_col(name: str) -> str:
    return " ".join(_norm_txt(name).split())


def _norm_match(value) -> str:
    s = _norm_txt(value)
    chunks = []
    for ch in s:
        chunks.append(ch if ch.isalnum() else " ")
    return " ".join("".join(chunks).split())


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

    def _find_cols(self, *candidates: str) -> list[str]:
        cols = []
        seen = set()
        for c in candidates:
            key = _norm_col(c)
            col = self.col_map.get(key)
            if col and col not in seen:
                seen.add(col)
                cols.append(col)
        return cols

    def _match_row_by_value(self, value: str, columns: list[str]):
        expected = _norm_match(value)
        if not expected:
            return None

        for col in columns:
            series = self.df[col].map(_norm_match)
            matched = self.df[series == expected]
            if not matched.empty:
                return matched.iloc[0]
        return None

    def _regional_lookup_columns(self, *, prefer_forti: bool = False) -> list[str]:
        primary = self._find_cols(
            "NOME_REGIONAL", "REGIONAL", "INTEGRADA", "REGIAO", "REGIÃO", "UF", "NOME_REGIONAL"
        )
        forti = self._find_cols("NOME_REG_FORTI")
        return forti + primary if prefer_forti else primary + forti

    def get_row_by_regional(self, regional: str, *, prefer_forti: bool = False):
        columns = self._regional_lookup_columns(prefer_forti=prefer_forti)
        if not columns:
            raise RuntimeError("Não encontrei coluna de REGIONAL na planilha. Me diga o nome exato da coluna.")
        return self._match_row_by_value(regional, columns)

    def get_emails_by_regional(self, regional: str) -> list[str]:
        # tenta achar colunas de email
        col_email_ger = self._find_col("EMAIL_GERENTE", "GERENTE_EMAIL", "EMAIL GERENTE", "E-MAIL GERENTE")
        col_email_dir = self._find_col("EMAIL_DIRETOR", "DIRETOR_EMAIL", "EMAIL DIRETOR", "E-MAIL DIRETOR")
        col_email_apo1 = self._find_col("EMAIL_APOIO_1", "EMAIL APOIO 1", "E-MAIL APOIO 1")
        col_email_apo2 = self._find_col("EMAIL_APOIO_2", "EMAIL APOIO 2", "E-MAIL APOIO 2")
        col_email_apo = self._find_col("EMAIL_APOIO", "APOIO_EMAIL", "EMAIL APOIO", "E-MAIL APOIO")
        col_email_unico = self._find_col("EMAIL", "E-MAIL", "MAIL")

        row = self.get_row_by_regional(regional)
        if row is None:
            return []
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

    def get_forti_name_by_regional(self, regional: str) -> str | None:
        row = self.get_row_by_regional(regional)
        if row is None:
            return None

        col_forti = self._find_col("NOME_REG_FORTI")
        if not col_forti:
            return None

        value = row.get(col_forti)
        if value is None:
            return None

        text = str(value).strip()
        return text or None

    def get_regional_name_for_forti(self, forti_name: str) -> str | None:
        row = self.get_row_by_regional(forti_name, prefer_forti=True)
        if row is None:
            return None

        col_reg = self._find_col("NOME_REGIONAL", "REGIONAL", "INTEGRADA", "REGIAO", "REGIÃO", "UF")
        if not col_reg:
            return None

        value = row.get(col_reg)
        if value is None:
            return None

        text = str(value).strip()
        return text or None