#include <SoftwareSerial.h>
const byte numChars = 32;
char receivedChars[numChars];   // an array to store the received data
static byte ndx = 0;

SoftwareSerial mySerial(10, 11); // RX, TX
boolean newData = false;

void setup() {
    Serial.begin(9600);
    mySerial.begin(2400);
}

void loop() {
    recvWithEndMarker();
    showNewData();
    
    
}

void recvWithEndMarker() {
    char endMarker = '=';
    char rc;
   
    if (mySerial.available() > 0 && newData == false) {
        rc = mySerial.read();
        if (rc != endMarker) {
            receivedChars[ndx] = rc;
            ndx++;
            if (ndx >= numChars) {
                ndx = numChars - 1;
            }
        }
        else {
            receivedChars[ndx] = '\0'; // terminate the string
           //ndx = 0;
            newData = true;
        }
    }
}

void showNewData() {
    if (newData == true) {
        strrevi(receivedChars);
        Serial.println(receivedChars);        
        //Serial.println(receivedChars[ndx]);
        newData = false;
        ndx=0;
       
        
    }
    
}


void strrevi(char *string)
{
    
    int len = 0;
    while (string[len])
        len++;

    int down = 0;
    int up = len - 1;

    while (up > down)
    {
        char c = string[down];
        string[down++] = string[up];
        string[up--] = c;
    }
   // delay(20);

}
