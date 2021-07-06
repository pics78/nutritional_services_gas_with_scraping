// コンテナおよび環境情報
// ====================================================
const COMMON_INFO = {
  'InactiveBackColor': '#efefef',
  'ActiveBackColor': null,
  'MAX_ROW': 900,
  'searchFoods': {
    'sheetName': '食品検索',
    'sourceInfo': {
      'url': 'https://fooddb.mext.go.jp/freeword/fword_select.pl',
      'method': 'post',
      'payloadTemp' : '{"SEARCH_WORD":"%s","function1":"検索"}',
    },
    'ranges': {
      /*検索ワード*/ 'searchWordCell': 'F6',
      /*ヒット数*/ 'hitCountCell': 'F7',
      /*食品番号表示列*/ 'foodNumberCol': 'A',
      /*食品名表示列*/ 'foodNameCol': 'B',
      /*検索結果表示開始行*/ 'resultFirstRow': 10,
    },
  },
  'calculateNutritionalValues': {
    'sheetName': '栄養価計算',
    'sourceInfo': {
      'urlTemp': 'https://fooddb.mext.go.jp/details/details.pl?ITEM_NO=1_%s_7',
    },
    'ranges': {
      /*食品番号列*/ 'foodNumberCol': 'A',
      /*食品名列*/ 'foodNameCol': 'B',
      /*重量列*/ 'weightCol': 'C',
      /*キャッシュ列*/ 'cacheCol': 'P',
      /*重量保持列*/ 'preWeightCol': 'Q',
      /*合計表示行*/ 'totalResultRow': 7,
      /*個別食品表示開始行*/ 'partialResultFirstRow': 11,
      /*総食品数表示セル*/ 'totalFoodCountCell': 'D2',
    },
    'nutInfo': {
      /*エネルギー*/ 'E': {
        'sheetCol': 'D',
        'sourceTableRow': 1,
        'digits': 0,
      },
      /*たんぱく質*/ 'P': {
        'sheetCol': 'E',
        'sourceTableRow': 4,
        'digits': 1,
      },
      /*脂質*/ 'F': {
        'sheetCol': 'F',
        'sourceTableRow': 6,
        'digits': 1,
      },
      /*炭水化物*/ 'Car': {
        'sheetCol': 'G',
        'sourceTableRow': 8,
        'digits': 1,
      },
      /*ナトリウム*/ 'Na': {
        'sheetCol': 'H',
        'sourceTableRow': 10,
        'digits': 0,
      },
      /*カリウム*/ 'K': {
        'sheetCol': 'I',
        'sourceTableRow': 11,
        'digits': 0,
      },
      /*カルシウム*/ 'Cal': {
        'sheetCol': 'J',
        'sourceTableRow': 12,
        'digits': 0,
      },
      /*鉄*/ 'Fe': {
        'sheetCol': 'K',
        'sourceTableRow': 15,
        'digits': 1,
      },
      /*ビタミンA*/ 'VA': {
        'sheetCol': 'L',
        'sourceTableRow': 28,
        'digits': 0,
      },
      /*ビタミンB1*/ 'VB1': {
        'sheetCol': 'M',
        'sourceTableRow': 35,
        'digits': 2,
      },
      /*ビタミンB2*/ 'VB2': {
        'sheetCol': 'N',
        'sourceTableRow': 36,
        'digits': 2,
      },
      /*ビタミンC*/ 'VC': {
        'sheetCol': 'O',
        'sourceTableRow': 44,
        'digits': 0,
      },
      /*食塩相当量*/ 'Sal': {
        'sheetCol': 'P',
        'sourceTableRow': 53,
        'digits': 1,
      },
    },
  },
};
// ====================================================