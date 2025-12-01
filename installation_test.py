#!/usr/bin/env python3
"""
Test script to verify the Employee Attendance Analytics installation
"""

import sys

def test_imports():
    """Test if all required packages are installed"""
    print("🔍 Testing package imports...\n")
    
    packages = {
        'streamlit': 'Streamlit',
        'pandas': 'Pandas',
        'numpy': 'NumPy',
        'openpyxl': 'OpenPyXL',
        'plotly': 'Plotly',
        'langchain': 'LangChain',
        'langchain_openai': 'LangChain OpenAI',
        'langchain_community': 'LangChain Community',
        'requests': 'Requests',
    }
    
    failed = []
    
    for package, name in packages.items():
        try:
            __import__(package)
            print(f"✅ {name:<20} - OK")
        except ImportError as e:
            print(f"❌ {name:<20} - FAILED")
            failed.append(package)
    
    print()
    
    if failed:
        print(f"⚠️  Failed to import: {', '.join(failed)}")
        print("   Run: pip install -r requirements.txt")
        return False
    else:
        print("✅ All packages installed successfully!")
        return True

def test_modules():
    """Test if custom modules load correctly"""
    print("\n🔍 Testing custom modules...\n")
    
    try:
        import attendance_analyzer
        print("✅ attendance_analyzer - OK")
    except Exception as e:
        print(f"❌ attendance_analyzer - FAILED: {e}")
        return False
    
    try:
        import llm_handler
        print("✅ llm_handler       - OK")
    except Exception as e:
        print(f"❌ llm_handler - FAILED: {e}")
        return False
    
    print("\n✅ All custom modules loaded successfully!")
    return True

def test_ollama():
    """Test if Ollama is available"""
    print("\n🔍 Testing Ollama availability...\n")
    
    try:
        import requests
        response = requests.get('http://localhost:11434/api/tags', timeout=2)
        if response.status_code == 200:
            models = response.json().get('models', [])
            print(f"✅ Ollama is running")
            if models:
                print(f"   Available models: {', '.join([m['name'] for m in models])}")
            else:
                print("   ⚠️  No models installed. Run: ollama pull llama2")
        else:
            print("⚠️  Ollama service not responding")
    except Exception:
        print("⚠️  Ollama not available (optional - you can use OpenAI instead)")

def main():
    """Run all tests"""
    print("=" * 60)
    print("  Employee Attendance Analytics - Installation Test")
    print("=" * 60)
    print()
    
    # Test imports
    imports_ok = test_imports()
    
    if not imports_ok:
        print("\n❌ Installation incomplete. Please install missing packages.")
        sys.exit(1)
    
    # Test modules
    modules_ok = test_modules()
    
    if not modules_ok:
        print("\n❌ Module loading failed. Check for errors.")
        sys.exit(1)
    
    # Test Ollama (optional)
    test_ollama()
    
    print("\n" + "=" * 60)
    print("✅ Installation test completed successfully!")
    print("=" * 60)
    print("\n🚀 You can now start the application:")
    print("   streamlit run app.py")
    print("   or")
    print("   ./run.sh (Linux/Mac)")
    print("   run.bat (Windows)")
    print()

if __name__ == "__main__":
    main()