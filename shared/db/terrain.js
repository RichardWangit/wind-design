/** shared/db/terrain.js  — 地況種類參數資料庫 */
// ─── Embedded database (normalized from SQLite) ───────────────────────────
const DB = {
  categories: [
    {
      id: "A",
      name: "地況 A",
      description: "大城市市中心區，至少有50%之建築物高度大於20公尺者。",
      application_condition: "建築物迎風向之前方至少800公尺或建築物高度10倍的範圍（兩者取大值）係屬此種條件下，才可使用地況A。",
      upwind_distance_m: 800,
      building_height_multiplier: 10
    },
    {
      id: "B",
      name: "地況 B",
      description: "大城市市郊、小市鎮或有許多像民舍高度（10～20公尺），或較民舍為高之障礙物分布其間之地區者。",
      application_condition: "建築物迎風向之前方至少500公尺或建築物高度10倍的範圍（兩者取大值）係屬此種條件下，方可使用地況B。",
      upwind_distance_m: 500,
      building_height_multiplier: 10
    },
    {
      id: "C",
      name: "地況 C",
      description: "平坦開闊之地面或草原或海岸或湖岸地區，其零星座落之障礙物高度小於10公尺者。",
      application_condition: "無特定前方距離要求，平坦開闊地面即可適用。",
      upwind_distance_m: null,
      building_height_multiplier: null
    }
  ],
  parameters: [
    { id:1,  terrain_id:"A", parameter_name:"風速垂直分布指數",  parameter_symbol:"α",    parameter_value:0.32,  unit:"無因次", description:"相對於10分鐘平均風速之垂直分布法則的指數" },
    { id:2,  terrain_id:"A", parameter_name:"梯度高度",          parameter_symbol:"zg",   parameter_value:500,   unit:"m",     description:"梯度高度，風速達到均勻分布之高度" },
    { id:3,  terrain_id:"A", parameter_name:"紊流強度係數",       parameter_symbol:"b̄",   parameter_value:0.45,  unit:"無因次", description:"用於計算紊流強度之係數 b̄" },
    { id:4,  terrain_id:"A", parameter_name:"紊流強度係數",       parameter_symbol:"c",    parameter_value:0.45,  unit:"無因次", description:"用於計算紊流強度之係數 c" },
    { id:5,  terrain_id:"A", parameter_name:"積分尺度參數",       parameter_symbol:"λ",    parameter_value:55,    unit:"m",     description:"積分尺度長度參數" },
    { id:6,  terrain_id:"A", parameter_name:"紊流強度",           parameter_symbol:"ε̄",   parameter_value:0.5,   unit:"無因次", description:"平均紊流強度" },
    { id:7,  terrain_id:"A", parameter_name:"最小計算高度",       parameter_symbol:"zmin", parameter_value:18,    unit:"m",     description:"風速垂直分布計算之最小高度" },

    { id:8,  terrain_id:"B", parameter_name:"風速垂直分布指數",  parameter_symbol:"α",    parameter_value:0.25,  unit:"無因次", description:"相對於10分鐘平均風速之垂直分布法則的指數" },
    { id:9,  terrain_id:"B", parameter_name:"梯度高度",          parameter_symbol:"zg",   parameter_value:400,   unit:"m",     description:"梯度高度，風速達到均勻分布之高度" },
    { id:10, terrain_id:"B", parameter_name:"紊流強度係數",       parameter_symbol:"b̄",   parameter_value:0.62,  unit:"無因次", description:"用於計算紊流強度之係數 b̄" },
    { id:11, terrain_id:"B", parameter_name:"紊流強度係數",       parameter_symbol:"c",    parameter_value:0.30,  unit:"無因次", description:"用於計算紊流強度之係數 c" },
    { id:12, terrain_id:"B", parameter_name:"積分尺度參數",       parameter_symbol:"λ",    parameter_value:98,    unit:"m",     description:"積分尺度長度參數" },
    { id:13, terrain_id:"B", parameter_name:"紊流強度",           parameter_symbol:"ε̄",   parameter_value:0.33,  unit:"無因次", description:"平均紊流強度" },
    { id:14, terrain_id:"B", parameter_name:"最小計算高度",       parameter_symbol:"zmin", parameter_value:9,     unit:"m",     description:"風速垂直分布計算之最小高度" },

    { id:15, terrain_id:"C", parameter_name:"風速垂直分布指數",  parameter_symbol:"α",    parameter_value:0.15,  unit:"無因次", description:"相對於10分鐘平均風速之垂直分布法則的指數" },
    { id:16, terrain_id:"C", parameter_name:"梯度高度",          parameter_symbol:"zg",   parameter_value:300,   unit:"m",     description:"梯度高度，風速達到均勻分布之高度" },
    { id:17, terrain_id:"C", parameter_name:"紊流強度係數",       parameter_symbol:"b̄",   parameter_value:0.94,  unit:"無因次", description:"用於計算紊流強度之係數 b̄" },
    { id:18, terrain_id:"C", parameter_name:"紊流強度係數",       parameter_symbol:"c",    parameter_value:0.20,  unit:"無因次", description:"用於計算紊流強度之係數 c" },
    { id:19, terrain_id:"C", parameter_name:"積分尺度參數",       parameter_symbol:"λ",    parameter_value:152,   unit:"m",     description:"積分尺度長度參數" },
    { id:20, terrain_id:"C", parameter_name:"紊流強度",           parameter_symbol:"ε̄",   parameter_value:0.20,  unit:"無因次", description:"平均紊流強度" },
    { id:21, terrain_id:"C", parameter_name:"最小計算高度",       parameter_symbol:"zmin", parameter_value:4.5,   unit:"m",     description:"風速垂直分布計算之最小高度" },
  ]
};

