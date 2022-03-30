#include <Wire.h>
#include <Adafruit_ADS1X15.h>

Adafruit_ADS1015 ADS;
Adafruit_ADS1015 ADSV;

void setup() {
  Serial.begin(2000000);
  ADS.begin();
  ADSV.begin(0x49);
}

void loop() {
  int16_t adcH0;
  int16_t adcH1;
  int16_t adcH2;
  int16_t adcH3;
  int16_t adcV0;
  int16_t adcV1;
  int16_t adcV2;
  int16_t adcV3;

  adcH0 = ADS.readADC_SingleEnded(0);
  adcH1 = ADS.readADC_SingleEnded(1);
  adcH2 = ADS.readADC_SingleEnded(2);
  adcH3 = ADS.readADC_SingleEnded(3);

  adcV0 = ADSV.readADC_SingleEnded(0);
  adcV1 = ADSV.readADC_SingleEnded(1);
  adcV2 = ADSV.readADC_SingleEnded(2);
  adcV3 = ADSV.readADC_SingleEnded(3);
  
  Serial.print(adcH0);
  Serial.print(" ");
  Serial.print(adcH1);
  Serial.print(" ");
  Serial.print(adcH2);
  Serial.print(" ");
  Serial.print(adcH3);
  Serial.print(" ");
  Serial.print(adcV0);
  Serial.print(" ");
  Serial.print(adcV1);
  Serial.print(" ");
  Serial.print(adcV2);
  Serial.print(" ");
  Serial.println(adcV3);
  

}
