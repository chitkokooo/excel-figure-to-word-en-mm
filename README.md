# excel-figure-to-word-en-mm
Microsoft Excel တွင် ဂဏန်းရိုက်လျှင် တန်ဖိုးကို စာဖြင့်ပြပေးသော custom function ‌အသုံးပြုပုံ ပြပေးထားတာပါ။

### လိုအပ်ချက်များ
- Excel File သည် macro enabled file ဖြစ်သော `.xlsm` file ဖြစ်ရမည်။
- Excel Options တွင် Macro File များကို run နိုင်ရန် ခွင့်ပြုပြီး ဖြစ်ရမည်။
- `hidden_var` ဟု အမည်ပေးထားသော Excel Sheet တခု တည်ဆောက်ပြီး ဖော်ပြပါ information များ ထည့်သွင်းထားရမည်။ တည်ဆောက်ပြီး Hide ထားလိုပါက Hide လုပ်ထားနိုင်သည်။
 
## hidden_var Excel Sheet
<img src="https://github.com/chitkokooo/excel-figure-to-word-en-mm/blob/main/hidden_var.png" alt="hidden_var excel Sheet">

### English to English
#### အသုံးပြုပုံ
```
  =ConvertFigureToWordEN(A1)
```
#### နမူနာ
```
123 = Kyats- One Hundred Twenty Three Only.
```

### English to Myanmar (Unicode Fonts)
#### အသုံးပြုပုံ
```
  =ConvertFigureToWordUnicode(A1)
```
#### နမူနာ
```
123 = ကျပ် တစ်ရာ နှစ်ဆယ် သုံး တိတိ
```

### English to Myanmar (ZawgyiOne Font)
#### အသုံးပြုပုံ
```
  =ConvertFigureToWordZawgyiOne(A1)
```
#### နမူနာ
```
123 = ကျပ် တစ်ရာ နှစ်ဆယ် သုံး တိတိ
```

`A1` သည် နမူနာ ပြထားသော Cell ဖြစ်ပြီး နှစ်သက်ရာ Cell/Number နှင့် ပြောင်းလဲ အသုံးပြုနိုင်သည်။


### နောက်ဆက်တွဲ တွေ့ရှိချက်များ
- `hidden_var` sheet ၏ မြန်မာစာ (ZawgyiOne & Pyidaungsu Columns) ၂ တိုင်ရှိ ***ဆယ်***၊ ***ရာ***၊ ***ထောင်***၊ ***သောင်း***၊ ***သိန်း*** cell ကွက်များတွင် spacebar ၁ ချက်စာစီ ပိုထည့်ပေးပါက Word Wrap ဖြစ်သော Columns ကျဉ်းသည့် Cell များ၌ ပိုမိုသေသပ်သော output ထုတ်ပေးနိုင်ကြောင်း တွေ့ရှိပါသည်။

## Resources
[Pyiduangsu Font](https://www.mmunicode.org/wiki/pyidaungsu-font/)

## Youtube Video Tutorial
[Excel Figure-to-Word Youtube Video](https://www.youtube.com/watch?v=uf49H02H7X4)

