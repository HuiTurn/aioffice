from __future__ import annotations

import unittest

from aioffice.core.ids import new_id


class StableIdTests(unittest.TestCase):
    def test_ids_are_prefixed_unique_and_ulid_sized(self) -> None:
        values = {new_id("para") for _ in range(100)}
        self.assertEqual(len(values), 100)
        self.assertTrue(all(value.startswith("para_") for value in values))
        self.assertTrue(all(len(value) == len("para_") + 26 for value in values))

    def test_invalid_prefix_is_rejected(self) -> None:
        with self.assertRaises(ValueError):
            new_id("../bad")


if __name__ == "__main__":
    unittest.main()
