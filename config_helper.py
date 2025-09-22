import pandas as pd
import os

def create_master_template(file_path="master_termsheets.xlsx"):
    """Create a master Excel template with all required columns"""
    
    # Define all columns based on your original requirements
    columns = [
        'Investment_Name', 'Issuer', 'Product_Name', 'Investment_Thematic', 'TYPE',
        'Coupon_QTR', 'Coupon_Rate_Annual', 'Product_Status', 'Knock_In_Percent', 
        'Knock_Out_Percent', 'Minimum_Tenor_Q', 'Maximum_Tenor_Q', 'Observation_Frequency',
        'CCY', 'Strike_Date', 'Issue_Date', 'ISIN',
        'Underlying_1', 'Underlying_2', 'Underlying_3', 'Underlying_4',
        'Investment_Amount', 'Total_Units', 'Notional_Value', 'AUD_Equivalent',
        'Maturity_Date', 'Revenue',
        'Underlying_1_Issue_Price', 'Underlying_2_Issue_Price', 'Underlying_3_Issue_Price', 'Underlying_4_Issue_Price',
        'Underlying_1_Knock_In_Price', 'Underlying_2_Knock_In_Price', 'Underlying_3_Knock_In_Price', 'Underlying_4_Knock_In_Price',
        'Underlying_1_Knock_Out', 'Underlying_2_Knock_Out', 'Underlying_3_Knock_Out', 'Underlying_4_Knock_Out',
        'Extraction_Date', 'Source_File'
    ]
    
    # Create empty DataFrame with columns
    df = pd.DataFrame(columns=columns)
    
    # Write to Excel
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Master_Data', index=False)
        
        # Add a reference sheet with column descriptions
        descriptions = pd.DataFrame({
            'Column': columns,
            'Description': [
                'Name of the investment product',
                'Issuing institution',
                'Product designation by issuer',
                'Investment theme/category',
                'Product type classification',
                'Quarterly coupon designation',
                'Annual coupon rate percentage',
                'Current status of product',
                'Knock-in barrier percentage',
                'Knock-out barrier percentage',
                'Minimum tenor in quarters',
                'Maximum tenor in quarters',
                'How often observations occur',
                'Currency denomination',
                'Strike price determination date',
                'Issue/launch date',
                'International Securities Identification Number',
                'First underlying asset',
                'Second underlying asset',
                'Third underlying asset',
                'Fourth underlying asset',
                'Total investment amount',
                'Number of units purchased',
                'Notional value of investment',
                'Australian Dollar equivalent',
                'Final maturity date',
                'Revenue generated',
                'Issue price of first underlying',
                'Issue price of second underlying',
                'Issue price of third underlying',
                'Issue price of fourth underlying',
                'Knock-in price of first underlying',
                'Knock-in price of second underlying',
                'Knock-in price of third underlying',
                'Knock-in price of fourth underlying',
                'Knock-out price of first underlying',
                'Knock-out price of second underlying',
                'Knock-out price of third underlying',
                'Knock-out price of fourth underlying',
                'Date data was extracted',
                'Source PDF filename'
            ]
        })
        
        descriptions.to_excel(writer, sheet_name='Column_Reference', index=False)
    
    print(f"Master template created: {file_path}")
    return file_path

def add_custom_issuer_pattern(issuer_key, identifiers, issuer_name, patterns):
    """
    Helper function to add custom issuer patterns
    
    Args:
        issuer_key: Unique key for the issuer (e.g., 'deutsche_bank')
        identifiers: List of text patterns that identify this issuer
        issuer_name: Full official name of the issuer
        patterns: Dictionary of regex patterns for extraction
    
    Example:
        add_custom_issuer_pattern(
            'deutsche_bank',
            ['DEUTSCHE BANK', 'DB Group'],
            'Deutsche Bank AG',
            {
                'isin': r'[A-Z]{2}[A-Z0-9]{10}',
                'coupon_rate': r'(\d+\.?\d*)\s*%.*(?:coupon|yield)',
                'knock_in': r'(\d+)%.*(?:barrier|protection)',
                'dates': r'(\d{1,2})\.(\d{1,2})\.(\d{4})'
            }
        )
    """
    
    new_pattern = {
        'identifiers': identifiers,
        'issuer_name': issuer_name,
        'patterns': patterns
    }
    
    print(f"Custom issuer pattern for {issuer_key}:")
    print(f"  Identifiers: {identifiers}")
    print(f"  Name: {issuer_name}")
    print(f"  Patterns: {list(patterns.keys())}")
    
    return {issuer_key: new_pattern}

def validate_master_file(file_path):
    """Validate that master file has correct structure"""
    
    if not os.path.exists(file_path):
        print(f"❌ File not found: {file_path}")
        return False
    
    try:
        df = pd.read_excel(file_path, sheet_name='Master_Data')
        required_columns = [
            'Investment_Name', 'Issuer', 'ISIN', 'Issue_Date', 
            'Maturity_Date', 'Currency', 'Notional_Value'
        ]
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"❌ Missing required columns: {missing_columns}")
            return False
        
        print(f"✅ Master file validated: {len(df)} existing rows")
        print(f"✅ All required columns present")
        return True
        
    except Exception as e:
        print(f"❌ Error validating file: {str(e)}")
        return False

# Example usage and testing
if __name__ == "__main__":
    # Create master template
    master_file = create_master_template()
    
    # Validate it
    validate_master_file(master_file)
    
    # Example of adding custom issuer
    custom_pattern = add_custom_issuer_pattern(
        'ubs',
        ['UBS AG', 'UBS Group', 'UBS Investment Bank'],
        'UBS AG',
        {
            'isin': r'[A-Z]{2}[A-Z0-9]{10}',
            'coupon_rate': r'(\d+\.?\d*)\s*%.*(?:coupon|rate|yield)',
            'knock_in': r'(\d+)%.*(?:knock.?in|barrier|protection)',
            'dates': r'(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(\d{4})'
        }
    )
    
    print("\nTo add this custom pattern to your extractor:")
    print("extractor.issuer_patterns.update(custom_pattern)")
