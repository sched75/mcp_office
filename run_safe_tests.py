#!/usr/bin/env python3
"""Run only tests that use proper mocks and don't launch real Office applications."""

import subprocess
import sys

def run_safe_tests():
    """Run tests that are known to work with mocks."""
    print("=== RUNNING SAFE TESTS (Outlook only) ===")
    
    # Only run Outlook tests which use proper mocks
    result = subprocess.run([
        sys.executable, "-m", "pytest", 
        "tests/test_outlook_service.py",
        "tests/test_outlook_extended.py",
        "-v", 
        "--tb=short",
        "--no-header"
    ], capture_output=True, text=True)
    
    print("STDOUT:")
    print(result.stdout)
    print("STDERR:")
    print(result.stderr)
    print(f"RETURN CODE: {result.returncode}")
    
    return result.returncode

if __name__ == "__main__":
    exit_code = run_safe_tests()
    sys.exit(exit_code)