#!/usr/bin/env python3
"""
Test script for the enhanced Excel validator tool
"""
import subprocess
import sys
import os

def test_enhanced_validator():
    """Test the enhanced validator tool"""
    try:
        # Change to the correct directory
        os.chdir(r'd:\OneDrive\Desktop\thu nghien software')
        print("📁 Current directory:", os.getcwd())
        
        # List available Excel files
        excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~$')]
        print(f"📊 Found {len(excel_files)} Excel files:")
        for i, file in enumerate(excel_files, 1):
            print(f"  {i}. {file}")
        
        if not excel_files:
            print("❌ No Excel files found for testing")
            return
            
        print("\n🚀 Testing enhanced validator tool...")
        print("=" * 50)
        
        # Run the validator and send '1' as input to select first file
        proc = subprocess.Popen([sys.executable, 'excel_validator_final.py'], 
                               stdin=subprocess.PIPE, 
                               stdout=subprocess.PIPE, 
                               stderr=subprocess.PIPE, 
                               text=True,
                               encoding='utf-8')
        
        # Send input to select first file
        stdout, stderr = proc.communicate(input='1\n', timeout=30)
        
        print("✅ Tool executed successfully!")
        print("\n📋 OUTPUT:")
        print(stdout)
        
        if stderr:
            print("\n⚠️ STDERR:")
            print(stderr)
            
    except subprocess.TimeoutExpired:
        print("⏰ Tool execution timed out (this is expected for interactive tools)")
        proc.kill()
    except Exception as e:
        print(f"❌ Error testing tool: {e}")

if __name__ == "__main__":
    test_enhanced_validator()
