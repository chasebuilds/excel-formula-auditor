from benchmark.generate import generate
from benchmark.run_benchmark import score


def test_generator_is_seeded() -> None:
    assert generate(1, 40)[1].planted == generate(1, 40)[1].planted
    assert generate(1, 40)[1].planted != generate(2, 40)[1].planted


def test_small_benchmark_run() -> None:
    s = score(3, 60)
    assert s.clean_flags == 0
    assert all(s.caught[i] == s.planted[i] for i in s.planted)
