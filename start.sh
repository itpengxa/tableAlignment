#!/bin/bash

# 财务数据处理工具 - 快速启动脚本

# 颜色定义
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${GREEN}=== 财务数据处理工具 ===${NC}"

# 检查Python版本
echo -e "${YELLOW}检查Python环境...${NC}"
python_version=$(python3 --version 2>&1 | awk '{print $2}')
echo "Python版本: $python_version"

# 检查依赖
echo -e "${YELLOW}检查依赖包...${NC}"
if ! python3 -c "import pandas" 2>/dev/null; then
    echo -e "${RED}缺少依赖包，正在安装...${NC}"
    pip install -r requirements.txt
else
    echo "依赖包已安装 ✓"
fi

# 启动应用
echo -e "${YELLOW}启动应用...${NC}"
echo -e "${GREEN}===================${NC}"
python3 main.py
