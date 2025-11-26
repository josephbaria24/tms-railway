# test_endpoints.py
# Run this to test if your endpoints are accessible
# Usage: python test_endpoints.py

import requests
import json

BASE_URL = "http://localhost:8000"

def test_endpoint(method, path, description):
    """Test a single endpoint"""
    url = f"{BASE_URL}{path}"
    print(f"\n{'='*60}")
    print(f"Testing: {description}")
    print(f"Method: {method}")
    print(f"URL: {url}")
    print(f"{'='*60}")
    
    try:
        if method == "GET":
            response = requests.get(url, timeout=5)
        elif method == "POST":
            response = requests.post(url, json={}, timeout=5)
        else:
            print(f"❌ Unknown method: {method}")
            return False
        
        print(f"Status: {response.status_code}")
        
        try:
            data = response.json()
            print(f"Response: {json.dumps(data, indent=2)}")
        except:
            print(f"Response: {response.text}")
        
        if response.status_code == 200:
            print(f"✅ SUCCESS")
            return True
        else:
            print(f"❌ FAILED")
            return False
            
    except Exception as e:
        print(f"❌ ERROR: {str(e)}")
        return False

def main():
    print("\n" + "="*60)
    print("TESTING FASTAPI ENDPOINTS")
    print("="*60)
    
    # Test all endpoints
    results = []
    
    results.append(test_endpoint("GET", "/", "Root endpoint"))
    results.append(test_endpoint("GET", "/health", "Health check"))
    results.append(test_endpoint("GET", "/database/stats", "Database stats"))
    results.append(test_endpoint("GET", "/database/backup", "Database backup"))
    results.append(test_endpoint("POST", "/database/reset", "Database reset"))
    results.append(test_endpoint("POST", "/database/delete-all-records", "Delete all records"))
    
    # Summary
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    passed = sum(results)
    total = len(results)
    print(f"Passed: {passed}/{total}")
    
    if passed == total:
        print("✅ All tests passed!")
    else:
        print(f"❌ {total - passed} test(s) failed")
    
    return passed == total

if __name__ == "__main__":
    main()