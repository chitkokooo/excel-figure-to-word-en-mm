# excel-figure-to-word-en-mm
Microsoft Excel တွင် ဂဏန်းရိုက်လျှင် တန်ဖိုးကို စာဖြင့်ပြပေးသော custom function ‌အသုံးပြုပုံ ပြပေးထားတာပါ။

### လိုအပ်ချက်များ
- Excel File သည် macro enabled file ဖြစ်သော `.xlsm` file ဖြစ်ရမည်။
- Excel Options တွင် Macro File များကို run နိုင်ရန် ခွင့်ပြုပြီး ဖြစ်ရမည်။
- `hidden_var` ဟု အမည်ပေးထားသော Excel Sheet တခု တည်ဆောက်ပြီး ဖော်ပြပါ information များ ထည့်သွင်းထားရမည်။ တည်ဆောက်ပြီး Hide လိုလုပ်က လုပ်ထားနိုင်သည်။

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
123 = ကျပ် တစ်ရာနှစ်ဆယ်သုံး တိတိ
```

### English to Myanmar (ZawgyiOne Font)
#### အသုံးပြုပုံ
```
  =ConvertFigureToWordZawgyiOne(A1)
```
#### နမူနာ
```
123 = ကျပ် တစ်ရာနှစ်ဆယ်သုံး တိတိ
```

`A1` သည် နမူနာ ပြထားသော Cell ဖြစ်ပြီး နှစ်သက်ရာ Cell/Number နှင့် ပြောင်းလဲ အသုံးပြုနိုင်သည်။

