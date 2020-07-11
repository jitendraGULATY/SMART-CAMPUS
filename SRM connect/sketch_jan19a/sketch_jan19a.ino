const int led=13;
#define fan 8
int value;
int val;
void setup() 
   { 
      Serial.begin(9600); 
      pinMode(led, OUTPUT);
      digitalWrite (led, LOW);
      Serial.println("Connection established...");
   }
 
void loop() 
   {
     while (Serial.available())
        {
           value = Serial.read();
        }
     
     if (value == '1')
        digitalWrite (led, HIGH);
     else if (value == '2')
        digitalWrite (fan, HIGH);
     else if (value == '3')
        digitalWrite (fan, LOW);   
     else if (value == '0')
        digitalWrite (led, LOW);
   }
