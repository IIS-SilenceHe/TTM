#$firstNames = (@'
#赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜戚谢邹喻柏水窦章
#云苏潘葛奚范彭郎鲁韦昌马苗凤花方俞任袁柳酆鲍史唐费廉岑薛雷贺倪汤滕殷罗毕郝邬安常
#乐于时傅皮卞齐康伍余元卜顾孟平黄和穆萧尹
#'@.Split(("`n", "`r")) -join $null).ToCharArray()

$firstNames = "何"

$secondNames = (@'
一丁三专世业丝中丰临丹丽举乃义乐乔书云亘亮人仁今仙令仪伟伯伶佑作佩佳侠侬俊俏俐信
修倩健偲儿允元兆光兰兴其典军冠冬冰凌凝凡凯刚初利力勃勇勋化北千卉华卓博卫卿厚原友
双发叡可叶司合吉同名向君吟听启周和咏咸品哲唱善喆喜嗣嘉囡国圣坚基堂墨壁夏多天奇奕
奥好如妃妍妙妞妮姗姝姣娅娇娜娟娥娴婉婧婵婷媚媛嫒嫔嫚子存孟季学宁宇安宜宝实宣宵家
宸容宾密寒寰寻寿小尘尚尹展山岑岚峯峰峻巍州工巧布帅帆希干平年广庆康庸延建弘强弼彤
彦彩彬彭影微德心志忠念忻怀思怡恩恬恺悟悦情惜惠愉意慈慕慧懋懿成才扬承抒拔振捷掣敏
教文斌斯新方施旎旭旻昂昊昌明易昕星春昶晋晓晔晖晤晨景晴晶智暄暖暮曜曦曲曼曾月朋朔
朗望朝木本朵杉杏材杰松林枝枫柏柔柳栋树格桃桐梅梓梦棠森楚楠欢欣歆歌正武毅民水永江
池沈沉沙沛河泉波泰泽洁洛津洮洲流济浓浦浩海涉涛润涵淑淳淼清渊温湃湉湘源溥溪滢滨漪
漫澄澍澹濮濯瀚灵灿炎炳烁烨焕焱然煜煦熙熠燕爽牧献玄玉玑玛玟玲珉珊珍珑珠珺琅琇琛琦
琨琪琲琳琴琼瑜瑞瑶瑾璇璞瓃甜用甫田甲男画畅略白皎益盼真睿知石碧磊礼祖祥祯祺禄福禧
禾秀秉秋程穆空立章童端竹笑笛筠简箫籁米精素红纬纳纵纶经绢绣绮维罗罡美羡羽翎翔翠翮
翰翼耀聪胜能自致舒舟航良艳艺艾芃芊芝芦芬花芳芸苑苗若英茂范茉茗茜茵荌荣荫莉莎莘莲
莹菁菊菡菱菲萌萍萝萧萱蒙蓉蓓蓝蔓蔚蕊蕙蕴蕾薄薇藉藻虎虹蝶行衣裕西言誉许识诗诚语诺
谊谧谷豪贝贞贤资赋赐赡赫超越跃路轩载辉辰达迈运进远迪逸邈邵郁郎采金鑫铃铄锋锐锦长
闲闵阳阵陶隽雄雅雨雪雯雰霁霓霖霞露青靓靖静韦音韵韶顺颀颖颜飇风飙飞香馨驰驹骄高魁
魄鲲鲸鸣鸾鸿鹍鹏麦默黛齐龙
'@.Split(("`n", "`r")) -join $null).ToCharArray()

$mobilePrefixes = @'
134 135 136 137 138 139 147 150 151 152 157 158 159 182 187 188 147 157 188 
130 131 132 155 156 185 186 186 133 153 180 189 189 200 133 150 151 152 153 
155 156 157 158 159 130 131 132 133 134 135 136 137 138 139 180 182 185 186 
187 188 189 170
'@.Split(("`n", "`r", " "), 'RemoveEmptyEntries')

$cityCodes = @'
410701 341101 130901 130601 331001 230501 370601 659001 450301 120221 620701 
341001 210201 510101 130501 320901 520401 510701 623001 450401 650201 420901 
360701 620401 510501 542521 620501 232701 530401 611001 340401 321301 520101 
320401 430501 653101 460101 220301 640201 610801 440501 450701 533421 429004 
522401 532901 532801 130101 652101 445301 320601 410301 140201 530501 632121 
511301 210501 152921 140901 330701 410101 350501 621101 350901 652801 652201 
210801 440801 341801 500101 371401 411301 341501 141101 371301 440301 522201 
231101 510901 330601 450201 350301 150201 220101 440901 411601 371101 320101 
632221 330301 622901 130201 140101 131001 430301 610701 640101 450801 360601 
131101 650101 440101 210301 632801 360401 231001 320801 530601 610201 210701 
632321 360501 440201 512001 542621 370901 320201 220501 230801 220401 110228 
430401 451201 340701 542301 130401 360201 410901 620601 150301 150801 140401 
451001 532501 370101 440601 321001 411701 321101 654002 451401 410501 220801 
350401 330201 522301 130701 420301 445101 420701 421301 211201 511701 420601 
'@.Split(("`n", "`r", " "), 'RemoveEmptyEntries')

function Get-RandomName {
    return ($firstNames|Get-Random) + 
    (($secondNames|Get-Random -Count ((1,2)|Get-Random)) -join $null)
}

function Get-RandomQQ {
    return [string](Get-Random -Minimum 100000 -Maximum 9999999999)
}

function Get-RandomEMail ($QQ) {
    if ($QQ) {
        return "$QQ@qq.com"
    } else {
        return (Get-RandomQQ) + '@qq.com'
    }
}

function Get-RandomMobile {
    return ($mobilePrefixes | Get-Random) +
    ((0..9 | Get-Random -count 8) -join $null)
}

function Get-RandomSex {
    return ('男', '女' | Get-Random)
}

function Get-RandomBirthday {
    return (Get-Date).AddDays(-(Get-Random -Maximum (365 * 110))).Date
}

function Get-RandomID  ([DateTime]$Birthday, $Sex){
    $cityCode = $cityCodes | Get-Random
    #$cityCode = (0..9 | Get-Random -Count 6) -join $null

    if (!$Birthday) {
        $Birthday = Get-RandomBirthday

    }
    $birthdayCode = '{0:yyyyMMdd}' -f $Birthday

    if ($Sex -eq '男' -or $Sex -eq $true) {
        $seq = (Get-Random -Minimum 0 -Maximum 49) * 2 + 1
    } else {
        $seq = (Get-Random -Minimum 0 -Maximum 49) * 2 + 2
    }

    $result = '{0}{1}{2:D3}' -f $cityCode, $birthdayCode, $seq
    $w = @(7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
    $total = 0
    for ($i = 0; $i -lt 17; $i++) {
        $total += $w[$i] * [int]::parse($result[$i])
    }
    $checkCode = [string]($total % 11)
    $checkCode = '10X98765432'[$checkCode]

    $result = $result + $checkCode
    return $result
}

function Get-RandomIdentity {
    $qq = Get-RandomQQ
    $birthday = Get-RandomBirthday
    $sex = Get-RandomSex
    return [pscustomobject][ordered]@{
        Name = Get-RandomName;
        QQ = $qq;
        EMail = Get-RandomEMail -QQ $qq;
        Mobile = Get-RandomMobile;
        Birthday = '{0:yyyy/MM/dd}' -f $birthday
        Sex = $sex
        ID = Get-RandomID -Birthday $birthday -Sex $sex
    }
}


1..100 | %{Get-RandomName;Write-Host}