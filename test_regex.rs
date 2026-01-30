use fancy_regex::Regex;

fn main() {
    // Test 1: hello.world should match "hello world"
    let pattern = "(?i)hello.world";
    let regex = Regex::new(pattern).unwrap();
    let text = "hello world";
    
    println!("Pattern: {}", pattern);
    println!("Text: '{}'", text);
    println!("Match: {:?}", regex.find(text).unwrap());
    
    // Test 2: with dot matches newline
    let pattern2 = "(?i)(?s)hello.world";
    let regex2 = Regex::new(pattern2).unwrap();
    let text2 = "hello\nworld";
    
    println!("\nPattern: {}", pattern2);
    println!("Text: 'hello\nworld'");
    println!("Match: {:?}", regex2.find(text2).unwrap());
    
    // Test 3: without (?s), dot should NOT match newline
    let pattern3 = "(?i)hello.world";
    let regex3 = Regex::new(pattern3).unwrap();
    let text3 = "hello\nworld";
    
    println!("\nPattern: {}", pattern3);
    println!("Text: 'hello\nworld'");
    println!("Match: {:?}", regex3.find(text3).unwrap());
}
