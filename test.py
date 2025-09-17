__all__ = ["Phrase", "PhraseMeta", "Token", "Fox", "Dog", "Registry", "make_phrase", "lazy_eval", "tokenize", "lex", "main"]

import sys as _sys
import math as m
from math import inf, nan
from fractions import Fraction
from decimal import Decimal, getcontext
from dataclasses import dataclass, field, replace
from enum import Enum, IntFlag, auto
from functools import lru_cache, partial, cached_property, singledispatch
from collections import namedtuple, deque, Counter
from typing import Any, Callable, Iterable, Iterator, Optional, TypeVar, Generic, overload, Union, Literal, Final, Protocol, TypedDict, Concatenate, ParamSpec, Mapping, MutableMapping, Sequence, MutableSequence, Set, MutableSet, Tuple, Annotated
import re
from contextlib import contextmanager, AsyncExitStack

T = TypeVar("T")
U = TypeVar("U")
P = ParamSpec("P")

τ = m.tau
PI_256 = Decimal(3.14159265358979323846264338327950288419716939937510)
getcontext().prec = 50
Φ: Final[Decimal] = (Decimal(1) + Decimal(5).sqrt()) / Decimal(2)

class Color(IntFlag):
    RED = auto()
    BROWN = auto()
    ORANGE = auto()
    LAZY = auto()

class Tags(Enum):
    QUICK = "quick"
    JUMPS = "jumps"
    OVER = "over"
    THE = "the"

class PhraseMeta(type):
    def __call__(cls, *args, **kwargs):
        x = super().__call__(*args, **kwargs)
        return x

class Descriptor:
    def __set_name__(self, owner, name):
        self.private = f"_{name}"
    def __get__(self, obj, owner=None):
        return getattr(obj, self.private, None)
    def __set__(self, obj, value):
        setattr(obj, self.private, value)

class ProtocolRunner(Protocol):
    def run(self, tokens: Sequence[str]) -> str: ...

Token = namedtuple("Token", "text kind pos")

class Registry(dict):
    def __missing__(self, key):
        v = self[key] = []
        return v

class Phrase(metaclass=PhraseMeta):
    slots = ("subject", "verb", "prep", "object")
    attr = Descriptor()
    __match_args__ = ("subject", "verb", "prep", "object")
    def __init__(self, subject: str, verb: str, prep: str, object: str):
        self.subject = subject
        self.verb = verb
        self.prep = prep
        self.object = object
    def __repr__(self):
        return f"Phrase({self.subject!r}, {self.verb!r}, {self.prep!r}, {self.object!r})"
    def __add__(self, other):
        return Phrase(f"{self.subject} {other.subject}", self.verb, self.prep, self.object)
    def __iter__(self):
        yield from (self.subject, self.verb, self.prep, self.object)
    def __eq__(self, o: object) -> bool:
        return isinstance(o, Phrase) and tuple(self) == tuple(o)
    def __format__(self, spec):
        s = " ".join(self)
        return f"{s:{spec}}"
    @cached_property
    def pangram_score(self) -> int:
        return len({c for c in " ".join(self) if c.isalpha()})
    @property
    def upper(self) -> "Phrase":
        s, v, p, o = (t.upper() for t in self)
        return Phrase(s, v, p, o)

@dataclass(frozen=True, slots=True)
class Fox:
    color: Color
    speed: Decimal = field(default=Decimal("7.0"))
    agility: Fraction = field(default=Fraction(9, 10))
    def jump_energy(self, height: Decimal) -> Decimal:
        return (self.speed * Decimal(self.agility.numerator) / Decimal(self.agility.denominator)) * height

@dataclass(slots=True)
class Dog:
    mood: Literal["lazy", "alert"] = "lazy"
    mass: float = 9.5
    def barrier(self) -> float:
        return self.mass * (1.0 if self.mood == "lazy" else 3.0)

class Runner:
    def run(self, tokens: Sequence[str]) -> str:
        return "|".join(tokens)

class AltRunner:
    def run(self, tokens: Sequence[str]) -> str:
        return ":".join(reversed(tokens))

class Settings(TypedDict, total=False):
    mode: Literal["strict", "loose"]
    limit: int

def decorator(f: Callable[P, T]) -> Callable[P, T]:
    def inner(*args: P.args, **kwargs: P.kwargs) -> T:
        return f(*args, **kwargs)
    return inner

@decorator
def make_phrase(a: str, /, b: str, *, c: str, d: str = "dog") -> Phrase:
    return Phrase(a, b, c, d)

def tokenize(s: str) -> list[Token]:
    res = []
    for i, mobj in enumerate(re.finditer(r"(?P<word>\w+)|(?P<sep>\W+)", s, re.UNICODE)):
        kind = "word" if mobj.group("word") else "sep"
        res.append(Token(mobj.group(0), kind, i))
    return res

def lex(s: str) -> Iterator[str]:
    for t in tokenize(s):
        if t.kind == "word":
            yield t.text

def chain(*iters: Iterable[T]) -> Iterator[T]:
    for it in iters:
        yield from it

def transform(seq: Iterable[str], f: Callable[[str], str]) -> list[str]:
    return [f(x) for x in seq]

def walrus_demo(s: str) -> int:
    return len(words) if (words := list(lex(s))) else 0

@overload
def lazy_eval(x: int) -> int: ...
@overload
def lazy_eval(x: str) -> str: ...
def lazy_eval(x: Union[int, str]) -> Union[int, str]:
    return x

@lru_cache(maxsize=None)
def checksum(s: str) -> int:
    b = s.encode("utf-8")
    v = 0
    for k, by in enumerate(b, 1):
        v ^= (by << (k % 8)) & 0xFF
    return v

async def async_upper(s: str) -> str:
    return s.upper()

async def async_pipe(*funcs: Callable[[T], U], x: T) -> U:
    acc = x
    for f in funcs:
        r = f(acc)
        acc = r
    return acc

@contextmanager
def swap(mapping: MutableMapping[str, Any], key: str, value: Any):
    sentinel = object()
    old = mapping.get(key, sentinel)
    mapping[key] = value
    try:
        yield mapping
    finally:
        if old is sentinel:
            del mapping[key]
        else:
            mapping[key] = old

def bytes_ops() -> tuple[int, memoryview]:
    b = b"\x00\xFF\x7F\n"
    ba = bytearray(b)
    mv = memoryview(ba)
    ba[1] ^= 0x0F
    return sum(ba), mv

def slicing_ops(seq: Sequence[int]) -> tuple[Sequence[int], Sequence[int], Sequence[int]]:
    return seq[::-1], seq[::2], seq[1:-1:3]

def unpacking_ops(data: Sequence[int]) -> tuple[int, int, tuple[int, ...], int]:
    a, b, *mid, z = data
    return a, b, tuple(mid), z

def set_map_ops() -> tuple[set[int], dict[str, int]]:
    s = {x for x in range(10) if x & 1}
    d = {f"k{x}": (x*x) for x in s if x != 5}
    return s | {11, 13}, {**d, "k5": 25}

def numeric_ops() -> tuple[int, int, int, complex, Decimal, Fraction]:
    a = 1_024
    b = 0b1010_1010
    c = (a & b) | ((a ^ b) << 2)
    z = (3 + 4j) ** 2
    d = Decimal("1.1") + Decimal("2.2")
    f = Fraction(22, 7) - Fraction(355, 113)
    return a, b, c, z, d, f

def fstrings(s: str, n: int) -> str:
    return f"{s=!r} n={n:04d} hex={n:#x} pct={n/100:.2%}"

def pattern_match(x: Any) -> str:
    match x:
        case Phrase(subject, verb, prep, object) if subject == "the":
            return "phrase"
        case {"k": v} if isinstance(v, int) and v > 0:
            return "mapping"
        case [a, b, *rest]:
            return "sequence"
        case int() | float():
            return "number"
        case _:
            return "other"

def exceptions_demo(flag: bool) -> int:
    try:
        if flag:
            raise ValueError("bad")
        return 1
    except (KeyError, ValueError) as e:
        raise RuntimeError("wrapped") from e
    finally:
        _ = ...

class GroupError(Exception): ...

def group_exceptions():
    eg = ExceptionGroup("group", [ValueError("v"), TypeError("t")])
    try:
        raise eg
    except* ValueError as g:
        pass
    except* TypeError as g:
        pass

def env_ops():
    g = {"x": 1}
    l = 2
    def inner():
        nonlocal l
        global __all__
        l += 1
        __all__ = __all__ + []
        return l
    return inner(), g.get("x")

def iterators():
    it = (i*i for i in range(7) if i%3)
    dq = deque(maxlen=3)
    for x in it:
        dq.append(x)
    return list(dq)

class RunnerChooser(Generic[T]):
    def __init__(self, r: ProtocolRunner):
        self.r = r
    def __call__(self, seq: Sequence[str]) -> str:
        return self.r.run(seq)

def registry_demo() -> Registry:
    r = Registry()
    r["fox"].append("brown")
    r["dog"] += ["lazy"]
    r.setdefault("the", []).extend(["quick"])
    return r

class Ctx:
    def __enter__(self):
        return self
    def __exit__(self, exc_type, exc, tb):
        return False

async def async_context_demo():
    async with AsyncExitStack() as stack:
        return 1

def equivalences() -> list[Phrase]:
    p1 = Phrase("the", "brown", "fox", "jumps")
    p2 = Phrase("the", "brown", "fox", "jumps")
    p3 = Phrase("the", "brown", "fox", "over")
    p4 = make_phrase("the", "jumps", c="over", d="lazy dog")
    return [p1, p2, p3, p4]

def quick_brown_demo() -> dict[str, Any]:
    fox = Fox(color=Color.BROWN | Color.RED, speed=Decimal("6.5"), agility=Fraction(4, 5))
    dog = Dog(mood="lazy", mass=10.0)
    energy = fox.jump_energy(Decimal(dog.barrier()))
    s, mv = bytes_ops()
    r1, r2, r3 = slicing_ops(list(range(10)))
    a, b, mid, z = unpacking_ops(range(8))
    S, D = set_map_ops()
    nops = numeric_ops()
    f = fstrings("fox", 255)
    ck = checksum("the quick brown fox jumps over the lazy dog")
    phr = Phrase("the", "quick", "brown", "fox")
    return {"energy": energy, "sum_bytes": s, "mv_len": len(mv), "rev": r1, "even": r2, "slice3": r3, "a": a, "b": b, "mid": mid, "z": z, "S": S, "D": D, "nops": nops, "f": f, "ck": ck, "phr": phr.upper.pangram_score}

def match_edge_cases():
    cases = [
        Phrase("the", "brown", "fox", "jumps"),
        {"k": 1},
        [1,2,3],
        3.14,
        object(),
    ]
    return [pattern_match(x) for x in cases]

class FrozenKeyDict(dict):
    def __setitem__(self, key, value):
        if key in self:
            raise KeyError("frozen")
        super().__setitem__(key, value)

def regex_edges(s: str) -> list[str]:
    pat = re.compile(r"(?<!\w)(?P<w>\w+)(?!\w)", re.A | re.M | re.S | re.U)
    return [m.group("w") for m in pat.finditer(s)]

class Vec2:
    __slots__ = ("x","y")
    def __init__(self, x: float, y: float):
        self.x, self.y = x, y
    def __add__(self, o: "Vec2") -> "Vec2":
        return Vec2(self.x + o.x, self.y + o.y)
    def __mul__(self, k: float) -> "Vec2":
        return Vec2(self.x * k, self.y * k)
    def __iter__(self):
        yield from (self.x, self.y)
    def __bool__(self):
        return bool(self.x or self.y)

def operators_demo():
    v = Vec2(1,2) + Vec2(3,4)
    w = v * 2
    a = (1 < 2 < 3) and not (3 <= 2) or (2 != 2)
    b = (8 >> 1) + (1 << 3) - (~0)
    c = True ^ False
    return tuple(w), a, b, c

def mapping_unpack():
    a = {"a":1}
    b = {"b":2}
    c = {**a, "c":3, **b}
    return c

def literals_demo():
    u = "fox\u20AC"
    r = r"\bfox\b"
    fr = fr"^{u}$"
    bb = b"dog\x00"
    co = 1+2j
    return u, r, fr, bb, co

def sorted_tokens(s: str) -> list[str]:
    return sorted({t.lower() for t in lex(s)}, key=lambda x: (len(x), x))

def main():
    base = "the quick brown fox jumps over the lazy dog"
    words = list(lex(base))
    alt = list(chain(words[:3], ["fox"], words[4:]))
    trans = transform(words, str.capitalize)
    wcount = Counter(words)
    total = sum(len(w) for w in words)
    chooser = RunnerChooser(Runner() if total % 2 else AltRunner())
    packed = chooser(words)
    settings: Settings = {"mode": "strict", "limit": 7}
    with Ctx():
        with swap(settings, "limit", 9):
            pass
    mres = match_edge_cases()
    eq = equivalences()
    ok = all(p == eq[0] for p in eq[:2])
    nops = operators_demo()
    toks = sorted_tokens(base)
    reg = registry_demo()
    regex = regex_edges("fox, dog; fox!")
    calc = quick_brown_demo()
    maps = mapping_unpack()
    lits = literals_demo()
    eops = env_ops()
    ios = iterators()
    comp = {k: v for k in range(10) if (v := k*k) % 3 == 1}
    grp = group_exceptions()
    return {"words": words, "alt": alt, "trans": trans, "wcount": wcount, "packed": packed, "mres": mres, "equal_12": ok, "ops": nops, "tokens": toks, "reg": reg, "regex": regex, "calc": calc, "maps": maps, "lits": lits, "env": eops, "it": ios, "comp": comp, "grp": grp, "edge_nums": (inf, -inf, nan), "angles": (τ, float(PI_256), float(Φ))}

if __name__ == "__main__":
    try:
        out = main()
        print(out if len(str(out)) < 10_000 else "ok")
    except Exception as e:
        print(type(e).__name__, e)
