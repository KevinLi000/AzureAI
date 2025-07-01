    def _is_complex_page(self, page):
        """检测页面是否包含复杂内容"""
        # 获取页面内容统计
        text = page.get_text()
        blocks = page.get_text("dict")["blocks"]
        
        # 增强图像检测 - 使用多种方法检测图像
        image_blocks = []
        
        # 方法1: 基于块类型检测图像
        basic_image_blocks = [b for b in blocks if b["type"] == 1]
        image_blocks.extend(basic_image_blocks)
        
        # 方法2: 使用get_images方法检测嵌入图像
        try:
            embedded_images = page.get_images()
            if embedded_images and len(embedded_images) > 0:
                # 将嵌入图像的信息添加到图像块列表
                for img in embedded_images:
                    # 检查这个图像是否已经在基本图像块中
                    xref = img[0]
                    already_detected = any(b.get("xref") == xref for b in basic_image_blocks)
                    
                    if not already_detected:
                        # 创建一个表示此图像的块
                        image_blocks.append({
                            "type": 1,  # 图像类型
                            "xref": xref,
                            "is_additional_image": True
                        })
        except Exception as e:
            print(f"使用get_images方法检测图像时出错: {e}")
        
        # 检查是否有图像
        has_images = len(image_blocks) > 0
        
        # 检查文本块数量
        text_blocks = [b for b in blocks if b["type"] == 0]
        many_text_blocks = len(text_blocks) > 15
        
        # 检查是否有表格
        # TableFinder对象不支持len()操作，改用其他方法检测表格
        try:
            # 使用表格特征检测是否存在表格
            tables_dict = page.find_tables().extract()
            has_tables = len(tables_dict) > 0
        except:
            # 备用方法：检查页面文本是否包含表格特征
            text_lower = text.lower()
            table_indicators = ['table', '表格', '列表', 'column', 'row', '行', '列']
            table_structure = text.count('|') > 5 or text.count('\t') > 5
            has_tables = any(indicator in text_lower for indicator in table_indicators) or table_structure
        
        # 检查是否有复杂布局
        has_complex_layout = False
        
        # 分析文本块的位置分布
        if len(text_blocks) > 5:
            # 收集所有文本块的x坐标
            x_positions = []
            for block in text_blocks:
                x_positions.append(block["bbox"][0])  # 左边界
            
            # 如果x坐标分布在多个不同位置，可能是多列布局
            x_pos_counter = Counter(int(x // 20) * 20 for x in x_positions)  # 按20点为间隔分组
            distinct_x_pos = len([k for k, v in x_pos_counter.items() if v > 2])  # 至少出现3次的x位置
            has_complex_layout = distinct_x_pos >= 3
        
        # 增强格式保留模式下更积极地判定为复杂页面
        if hasattr(self, 'format_preservation_level'):
            if self.format_preservation_level == "maximum":
                # 最大保留模式 - 只要有任何一个复杂因素就判定为复杂
                return has_images or has_tables or has_complex_layout or many_text_blocks
            elif self.format_preservation_level == "enhanced":
                # 增强保留模式 - 至少两个复杂因素
                complexity_factors = sum([has_images, has_tables, has_complex_layout, many_text_blocks])
                return complexity_factors >= 2
        
        # 标准判定 - 如果页面有很多图像或表格，或者文本块非常多，则认为是复杂页面
        return (has_images and has_tables) or (has_images and many_text_blocks) or (has_tables and many_text_blocks)
    
    def _add_border_to_picture(self, pic_element, width=1, color="000000"):
        """为Word文档中的图片添加边框"""
        try:
            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn
            
            # 创建边框元素
            border = OxmlElement('w:pict')
            border.set(qn('xmlns:w'), 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            
            # 设置边框属性
            line_element = OxmlElement('v:line')
            line_element.set('from', '0,0')
            line_element.set('to', '100%,100%')
            
            stroke_element = OxmlElement('v:stroke')
            stroke_element.set('weight', f"{width}pt")
            stroke_element.set('color', color)
            
            line_element.append(stroke_element)
            border.append(line_element)
            
            # 将边框添加到图片元素
            pic_element.append(border)
        except Exception as e:
            print(f"添加图片边框失败: {e}")
            import traceback
            traceback.print_exc()
            
        return pic_element
