#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
行政处罚文档生成工具
功能：读取Excel数据，基于Word模板生成文档
作者：AI Assistant
日期：2024
"""

import os
import sys
import json
import logging
from datetime import datetime
from pathlib import Path

# 尝试导入必要的库
try:
    from docxtpl import DocxTemplate
    import pandas as pd
except ImportError as e:
    print(f"错误：缺少必要的库 - {e}")
    print("请运行: pip install docxtpl pandas openpyxl")
    input("按回车键退出...")
    sys.exit(1)


# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('生成日志.txt', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class DocumentGenerator:
    """文档生成器类"""
    
    def __init__(self, config_path='config.json'):
        """初始化生成器"""
        self.config = self.load_config(config_path)
        self.base_dir = Path(__file__).parent
        self.output_dir = self.base_dir / self.config.get('output_dir', '输出文件')
        self.ensure_directories()
        
    def load_config(self, config_path):
        """加载配置文件"""
        default_config = {
            "template_file": "模板文件/行政处罚.docx",
            "data_file": "信息文件/数据.xlsx",
            "output_dir": "输出文件",
            "filename_template": "{案号}_{当事人名称}_行政处罚决定书.docx",
            "sheet_name": 0,
            "skip_rows": 0,
            "encoding": "utf-8"
        }
        
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                    default_config.update(user_config)
                    logger.info(f"已加载配置文件: {config_path}")
            except Exception as e:
                logger.warning(f"加载配置文件失败: {e}，使用默认配置")
        else:
            logger.info("未找到配置文件，使用默认配置")
            # 创建默认配置文件
            self.create_default_config(config_path, default_config)
            
        return default_config
    
    def create_default_config(self, config_path, config):
        """创建默认配置文件"""
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            logger.info(f"已创建默认配置文件: {config_path}")
        except Exception as e:
            logger.warning(f"创建配置文件失败: {e}")
    
    def ensure_directories(self):
        """确保必要的目录存在"""
        dirs = [
            self.base_dir / '模板文件',
            self.base_dir / '信息文件',
            self.output_dir
        ]
        for d in dirs:
            d.mkdir(exist_ok=True)
            
    def validate_template(self, template_path):
        """验证模板文件"""
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"模板文件不存在: {template_path}")
        
        # 尝试打开模板
        try:
            doc = DocxTemplate(template_path)
            logger.info(f"模板文件验证通过: {template_path}")
            return True
        except Exception as e:
            raise Exception(f"模板文件验证失败: {e}")
    
    def read_data(self, data_path):
        """读取Excel数据"""
        if not os.path.exists(data_path):
            raise FileNotFoundError(f"数据文件不存在: {data_path}")
        
        try:
            # 转换为字符串路径
            data_path_str = str(data_path)
            
            # 支持 .xlsx 和 .xls 格式
            if data_path_str.endswith('.xls'):
                df = pd.read_excel(
                    data_path_str,
                    sheet_name=self.config['sheet_name'],
                    skiprows=self.config['skip_rows']
                )
            else:
                df = pd.read_excel(
                    data_path_str,
                    sheet_name=self.config['sheet_name'],
                    skiprows=self.config['skip_rows'],
                    engine='openpyxl'
                )
            
            # 清理数据：去除空行，处理NaN值
            df = df.dropna(how='all')  # 删除全空行
            df = df.fillna('')  # NaN转为空字符串
            
            # 清理列名（去除空格）
            df.columns = [str(col).strip() for col in df.columns]
            
            logger.info(f"成功读取数据文件: {len(df)} 条记录")
            return df
            
        except Exception as e:
            raise Exception(f"读取数据文件失败: {e}")
    
    def generate_filename(self, row_data):
        """根据配置生成文件名"""
        try:
            filename = self.config['filename_template'].format(**row_data)
            # 清理非法字符
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                filename = filename.replace(char, '_')
            return filename
        except KeyError as e:
            # 如果模板中的字段不存在，使用默认文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            return f"行政处罚决定书_{timestamp}_{id(row_data)}.docx"
        except Exception as e:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            return f"行政处罚决定书_{timestamp}.docx"
    
    def generate_document(self, template_path, row_data, output_path):
        """生成单个文档"""
        try:
            doc = DocxTemplate(template_path)
            
            # 添加系统变量
            context = dict(row_data)
            context['生成日期'] = datetime.now().strftime("%Y年%m月%d日")
            context['生成时间'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            context['年份'] = datetime.now().year
            
            # 渲染文档
            doc.render(context)
            
            # 保存文档
            doc.save(output_path)
            return True
            
        except Exception as e:
            logger.error(f"生成文档失败: {e}")
            return False
    
    def run(self):
        """运行生成流程"""
        logger.info("=" * 50)
        logger.info("开始生成行政处罚文档")
        logger.info("=" * 50)
        
        try:
            # 路径设置
            template_path = self.base_dir / self.config['template_file']
            data_path = self.base_dir / self.config['data_file']
            
            # 验证模板
            logger.info(f"验证模板文件: {template_path}")
            self.validate_template(template_path)
            
            # 读取数据
            logger.info(f"读取数据文件: {data_path}")
            df = self.read_data(data_path)
            
            if len(df) == 0:
                logger.warning("数据文件为空，没有生成任何文档")
                return
            
            # 生成文档
            success_count = 0
            fail_count = 0
            
            for index, row in df.iterrows():
                # 将行数据转为字典
                row_data = row.to_dict()
                
                # 生成文件名
                filename = self.generate_filename(row_data)
                output_path = self.output_dir / filename
                
                logger.info(f"[{index + 1}/{len(df)}] 正在生成: {filename}")
                
                # 生成文档
                if self.generate_document(template_path, row_data, output_path):
                    success_count += 1
                    logger.info(f"  ✓ 成功生成: {output_path}")
                else:
                    fail_count += 1
                    logger.error(f"  ✗ 生成失败: {filename}")
            
            # 输出统计
            logger.info("=" * 50)
            logger.info(f"生成完成！成功: {success_count}，失败: {fail_count}")
            logger.info(f"输出目录: {self.output_dir}")
            logger.info("=" * 50)
            
            print(f"\n生成完成！")
            print(f"成功: {success_count} 个文档")
            print(f"失败: {fail_count} 个文档")
            print(f"输出位置: {self.output_dir}")
            
        except FileNotFoundError as e:
            logger.error(f"文件错误: {e}")
            print(f"\n错误: {e}")
            print("请检查文件路径是否正确")
        except Exception as e:
            logger.error(f"程序运行错误: {e}", exc_info=True)
            print(f"\n错误: {e}")


def main():
    """主函数"""
    print("=" * 50)
    print("行政处罚文档生成工具")
    print("=" * 50)
    print()
    
    # 检查命令行参数
    config_file = sys.argv[1] if len(sys.argv) > 1 else 'config.json'
    
    # 创建生成器并运行
    generator = DocumentGenerator(config_file)
    generator.run()
    
    print()
    input("按回车键退出...")


if __name__ == '__main__':
    main()
