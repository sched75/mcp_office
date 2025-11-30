#!/usr/bin/env python3
"""Simple test runner to check if tests can execute."""

import subprocess
import sys

def run_tests():
    """Run pytest with detailed output."""
    print("=== RUNNING TESTS WITH DETAILED OUTPUT ===")
    
    # Try running tests with detailed output
    result = subprocess.run([
        sys.executable, "-m", "pytest", 
        "tests/", 
        "-v", 
        "--tb=short",
        "--no-header",
        "--no-summary"
    ], capture_output=True, text=True)
    
    print("STDOUT:")
    print(result.stdout)
    print("STDERR:")
    print(result.stderr)
    print(f"RETURN CODE: {result.returncode}")
    
    return result.returncode

if __name__ == "__main__":
    exit_code = run_tests()
    sys.exit(exit_code)