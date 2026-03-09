from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime, Text, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from datetime import datetime

DATABASE_URL = "sqlite:///./bizdb.db"
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# 書類種別
DOC_TYPES = ["見積書", "請求書", "納品書", "発注書", "領収書"]


class Document(Base):
    __tablename__ = "documents"
    id           = Column(Integer, primary_key=True, index=True)
    doc_type     = Column(String, default="見積書")   # 書類種別
    doc_number   = Column(String, default="")          # 書類番号
    filename     = Column(String, default="")
    client_name  = Column(String, default="")          # 取引先
    doc_date     = Column(String, default="")          # 書類日付
    due_date     = Column(String, default="")          # 支払期限 / 納期
    total_amount = Column(Float,  default=0)
    tax_amount   = Column(Float,  default=0)
    currency     = Column(String, default="JPY")
    status       = Column(String, default="未処理")    # 未処理/処理済/キャンセル
    memo         = Column(Text,   default="")
    raw_text     = Column(Text,   default="")
    uploaded_at  = Column(DateTime, default=datetime.now)
    items        = relationship("Item", back_populates="document",
                                cascade="all, delete-orphan")


class Item(Base):
    __tablename__ = "items"
    id           = Column(Integer, primary_key=True, index=True)
    document_id  = Column(Integer, ForeignKey("documents.id"))
    name         = Column(String, default="")
    quantity     = Column(Float,  default=1)
    unit         = Column(String, default="")
    unit_price   = Column(Float,  default=0)
    total_price  = Column(Float,  default=0)
    note         = Column(String, default="")
    # 相場情報
    market_low   = Column(Float,  default=0)   # 相場下限
    market_high  = Column(Float,  default=0)   # 相場上限
    market_memo  = Column(String, default="")  # 相場備考
    document     = relationship("Document", back_populates="items")


Base.metadata.create_all(bind=engine)


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
