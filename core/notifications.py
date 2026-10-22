from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Protocol


@dataclass(frozen=True)
class EventoNotificacao:
    tipo: str
    assunto: str
    mensagem: str
    equipamento: str | None = None
    ocorrido_em: datetime | None = None


class CanalNotificacao(Protocol):
    nome: str

    def enviar(self, evento: EventoNotificacao, destinatarios: list[str]) -> None:
        """Envia um evento por um canal externo, quando o canal for implementado."""


class CanalEmail(CanalNotificacao, Protocol):
    pass


class CanalTeams(CanalNotificacao, Protocol):
    pass


class CanalMicrosoft365(CanalNotificacao, Protocol):
    pass
